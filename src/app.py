from __future__ import annotations

from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
from datetime import date, datetime, timedelta, timezone
import hashlib
import math
from io import BytesIO
import json
import os
from pathlib import Path
import shutil
import time
from typing import Callable, Optional
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
from sqlalchemy import case, func, text
from sqlmodel import Session, select
from starlette.middleware.base import BaseHTTPMiddleware

from .db import DB_PATH, engine, get_session, init_db
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
    purge_retired_employees,
    role_sort_key,
    normalize_role,
)
from .models import CliMatrixOverdueSnapshot, CliMatrixSummarySnapshot, Employee, Requirement, SstsDeviceSnapshot, SstsSnapshotRun
from .seed import seed_all
from processor import build_sheet2_df, build_summary_df

BASE_PATH = Path(__file__).resolve().parent.parent
GOOGLE_EMPLOYEE_STATION_TABS = ["North", "South", "KOAA", "DDJ", "RHA", "NH", "BT"]
HIDDEN_EMPLOYEE_ROLES = {normalize_role("Chief Loco Inspector")}
EMPLOYEE_SYNC_BACKUP_DIR = DB_PATH.parent / "employee_sync_backups"
EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE = DB_PATH.parent / "employee_master_source_snapshot.json"
EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE = DB_PATH.parent / "employee_master_service_snapshot.json"
EMPLOYEE_MASTER_MISMATCH_ACTIONS_FILE = DB_PATH.parent / "employee_master_mismatch_actions.json"
EMPLOYEE_MASTER_REVIEW_REPORT_FILE = DB_PATH.parent / "employee_master_review_report.json"
EMPLOYEE_MASTER_KEEP_BOTH_FILE = DB_PATH.parent / "employee_master_keep_both.json"
EMPLOYEE_MASTER_EXTRA_REVIEW_KEEP_FILE = DB_PATH.parent / "employee_master_extra_review_keep.json"
CLI_NOMINATION_MISMATCH_ACTIONS_FILE = DB_PATH.parent / "cli_nomination_mismatch_actions.json"
LI_GRADING_METADATA_FILE = DB_PATH.parent / "li_grading_metadata.json"
GOOGLE_SHEETS_READONLY_SCOPE = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
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
ASSET_VER = "v20260514c"
templates.env.globals["asset_ver"] = ASSET_VER
TOP_PERFORMER_STORE_PATH = BASE_PATH / "data" / "top_performer_store.json"
TOP_PERFORMER_PHOTO_DIR = BASE_PATH / "static" / "top_performer_photos"
SSTS_JUNK_FILE_PATTERNS = [
    BASE_PATH / "__pycache__",
    BASE_PATH / "src" / "__pycache__",
    BASE_PATH / "tmp_upload.xlsx",
    BASE_PATH / "temp_openapi.json",
]


CLI_NAME_MANUAL_ALIASES = {
    "ATKHAN": "ABU TAYAB KHAN",
    "SAMARESHMONDAL": "SAMARESH MANDAL",
    "SANJAYKRGUPTA": "SANJAY KUMAR GUPTA",
    "SHYAMALPRADHAN": "SHYAMAL KUMAR PRADHAN",
    "SHIVASANKARMONDAL": "SHIVA SHANKAR MANDAL",
    "SIDDHARTHABISWAS": "SIDHARTHA BISWAS",
    "SUMITBHATTACHARJEE": "SUMIT BHATTACHERJEE",
    "SUSOVANKAR": "SUSHOVAN KAR",
    "TAMOJITNANDY": "TAMOJIT NANDI",
    "TAPASKRDE": "TAPAS KUMAR DE I",
}


def _normalize_cli_tokens(value: str) -> list[str]:
    text = str(value or "").strip()
    text = re.sub(r"\s*\([A-Za-z]{2,}[A-Za-z0-9-]*\d+[A-Za-z0-9-]*\)\s*$", "", text)
    cleaned = re.sub(r"[^A-Za-z]+", " ", text).strip()
    return [token for token in cleaned.upper().split() if token]


def _cli_names_equivalent(a: Optional[str], b: Optional[str]) -> bool:
    if not a or not b:
        return False
    if a.strip().casefold() == b.strip().casefold():
        return True
    tokens_a = _normalize_cli_tokens(a)
    tokens_b = _normalize_cli_tokens(b)
    if not tokens_a or not tokens_b:
        return False
    if tokens_a[-1] != tokens_b[-1]:
        return False
    shorter, longer = (tokens_a, tokens_b) if len(tokens_a) <= len(tokens_b) else (tokens_b, tokens_a)
    for idx, token in enumerate(shorter[:-1]):
        if idx >= len(longer) - 1:
            return False
        target = longer[idx]
        if token == target:
            continue
        if len(token) == 1 and target.startswith(token):
            continue
        return False
    return True


def _ensure_employee_sync_backup_dir() -> Path:
    EMPLOYEE_SYNC_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    return EMPLOYEE_SYNC_BACKUP_DIR


def _iter_ssts_cleanup_targets() -> list[Path]:
    targets: list[Path] = []
    targets.extend(SSTS_JUNK_FILE_PATTERNS)
    backup_dir = DB_PATH.parent
    if backup_dir.exists():
        targets.extend(sorted(backup_dir.glob("hr_backup_*.db")))
    employee_backup_dir = EMPLOYEE_SYNC_BACKUP_DIR
    if employee_backup_dir.exists():
        targets.extend(sorted(employee_backup_dir.glob("*.db")))
    unique_targets: list[Path] = []
    seen: set[str] = set()
    for path in targets:
        key = str(path.resolve()) if path.exists() else str(path)
        if key in seen:
            continue
        seen.add(key)
        unique_targets.append(path)
    return unique_targets


def _path_size_bytes(path: Path) -> int:
    if not path.exists():
        return 0
    if path.is_file():
        return path.stat().st_size
    total = 0
    for child in path.rglob("*"):
        if child.is_file():
            try:
                total += child.stat().st_size
            except OSError:
                continue
    return total


def _format_storage_size(size_bytes: int) -> str:
    if size_bytes <= 0:
        return "0 B"
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def _ssts_cleanup_junk_summary() -> dict[str, object]:
    items: list[dict[str, object]] = []
    total_bytes = 0
    for path in _iter_ssts_cleanup_targets():
        if not path.exists():
            continue
        size_bytes = _path_size_bytes(path)
        total_bytes += size_bytes
        items.append(
            {
                "label": path.relative_to(BASE_PATH).as_posix() if path.exists() and BASE_PATH in path.resolve().parents else path.name,
                "size_bytes": size_bytes,
                "size_label": _format_storage_size(size_bytes),
                "kind": "folder" if path.is_dir() else "file",
            }
        )
    return {
        "items": items,
        "item_count": len(items),
        "total_bytes": total_bytes,
        "total_label": _format_storage_size(total_bytes),
    }


def _delete_ssts_junk_files() -> dict[str, object]:
    deleted_items: list[str] = []
    freed_bytes = 0
    for path in _iter_ssts_cleanup_targets():
        if not path.exists():
            continue
        size_bytes = _path_size_bytes(path)
        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()
        except OSError:
            continue
        deleted_items.append(path.relative_to(BASE_PATH).as_posix() if BASE_PATH in path.resolve().parents else path.name)
        freed_bytes += size_bytes
    return {
        "deleted_items": deleted_items,
        "deleted_count": len(deleted_items),
        "freed_bytes": freed_bytes,
        "freed_label": _format_storage_size(freed_bytes),
    }


def _create_employee_sync_backup(tag: str = "pre_sync") -> str:
    backup_dir = _ensure_employee_sync_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"employee_sync_{tag}_{timestamp}.db"
    target = backup_dir / filename
    shutil.copy2(DB_PATH, target)
    return filename


def _latest_employee_sync_backup() -> tuple[Optional[Path], str]:
    backup_dir = _ensure_employee_sync_backup_dir()
    backups = sorted(backup_dir.glob("employee_sync_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not backups:
        return None, ""
    latest = backups[0]
    label = latest.name.replace("employee_sync_", "").replace(".db", "").replace("_", " ")
    return latest, label


def infer_report_date(filename: str | None) -> date | None:
    text = str(filename or "")
    for pattern in (r"(\d{2})[-_ ](\d{2})[-_ ](\d{4})", r"(\d{4})[-_ ](\d{2})[-_ ](\d{2})"):
        match = re.search(pattern, text)
        if not match:
            continue
        parts = [int(part) for part in match.groups()]
        try:
            if len(str(parts[0])) == 4:
                return date(parts[0], parts[1], parts[2])
            return date(parts[2], parts[1], parts[0])
        except ValueError:
            return None
    return None


def coerce_report_date(value: object | None) -> date | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _save_li_grading_metadata(filename: str | None) -> None:
    report_date = infer_report_date(filename or "")
    payload = {
        "filename": filename or "",
        "report_date": report_date.isoformat() if report_date else "",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    LI_GRADING_METADATA_FILE.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")


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


def _save_cli_bio_reference_rows(
    session: Session,
    records: list[dict[str, object]],
    source_filename: str,
) -> None:
    unique_rows: dict[str, dict[str, str]] = {}
    for record in records:
        cli_id = str(record.get("cli_id") or "").strip()
        cli_name = str(record.get("cli_name") or record.get("name") or "").strip()
        mobile_no = str(record.get("mobile_no") or record.get("mobile") or record.get("phone") or "").strip()
        if not cli_id or not cli_name:
            continue
        unique_rows[cli_id] = {"cli_name": cli_name, "mobile_no": mobile_no}

    session.execute(text("DELETE FROM cli_bio_reference;"))
    for cli_id, row in sorted(unique_rows.items()):
        session.execute(
            text(
                """
                INSERT INTO cli_bio_reference (cli_id, cli_name, mobile_no, gradation, source_file, updated_at)
                VALUES (:cli_id, :cli_name, :mobile_no, '0', :source_file, CURRENT_TIMESTAMP)
                """
            ),
            {
                "cli_id": cli_id,
                "cli_name": row["cli_name"],
                "mobile_no": row["mobile_no"] or None,
                "source_file": source_filename,
            },
        )


def _load_cli_bio_reference_rows(session: Session) -> list[dict[str, str]]:
    rows = session.exec(
        text(
            """
            SELECT cli_id, cli_name, COALESCE(mobile_no, ''), gradation, source_file, updated_at
            FROM cli_bio_reference
            ORDER BY cli_name, cli_id
            """
        )
    ).fetchall()
    return [
        {
            "cli_id": str(row[0] or ""),
            "cli_name": str(row[1] or ""),
            "mobile_no": str(row[2] or ""),
            "gradation": str(row[3] or "0"),
            "source_file": str(row[4] or ""),
            "updated_at": str(row[5] or ""),
        }
        for row in rows
    ]


def _refresh_cli_bio_reference_from_records(
    session: Session,
    records: list[dict[str, object]],
    source_filename: str,
) -> None:
    _save_cli_bio_reference_rows(session, records, source_filename)


def _cli_matrix_record_date(value: object) -> date | None:
    if value is None:
        return None
    if str(value).strip().lower() == "nat":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        if hasattr(value, "to_pydatetime"):
            converted = value.to_pydatetime()
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
            return None
    except Exception:
        return None
    return None


def _save_cli_matrix_snapshots(
    session: Session,
    report_date_value: date,
    summary_df,
    overdue_df,
) -> None:
    existing_summary = session.exec(
        select(CliMatrixSummarySnapshot).where(CliMatrixSummarySnapshot.report_date == report_date_value)
    ).all()
    for row in existing_summary:
        session.delete(row)

    existing_overdue = session.exec(
        select(CliMatrixOverdueSnapshot).where(CliMatrixOverdueSnapshot.report_date == report_date_value)
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
                oldest_fp_overdue_date=_cli_matrix_record_date(record["Oldest FP OverDue Date"]),
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
                oldest_fp_overdue_date=_cli_matrix_record_date(record["Oldest FP OverDue Date"]),
                counsel_over_due=int(record["Counsel Over Due"]),
                oldest_counsel_overdue_date=_cli_matrix_record_date(record["Oldest Counsel OverDue Date"]),
                grading_overdue=int(record["Grading OverDue"]),
                oldest_grading_overdue_date=_cli_matrix_record_date(record["Oldest Grading OverDue"]),
                total_over_due_cases=int(record["Total Over Due Cases"]),
            )
        )

    session.commit()


def _split_cli_name_and_inline_id(value: object | None) -> tuple[str, str]:
    text = " ".join(str(value or "").strip().split())
    if not text:
        return "", ""
    match = re.match(
        r"^(?P<name>.*?)(?:\s*\((?P<id>[A-Za-z]{2,}[A-Za-z0-9-]*\d+[A-Za-z0-9-]*)\))\s*$",
        text,
    )
    if not match:
        return text, ""
    name_text = " ".join((match.group("name") or "").split())
    inline_id = " ".join((match.group("id") or "").split())
    if not name_text:
        return text, ""
    return name_text, inline_id


def _clean_cli_id(value: object | None) -> str:
    text = " ".join(str(value or "").strip().split())
    return text.upper() if text else ""


def _clean_cli_name(value: object | None) -> str:
    text, _ = _split_cli_name_and_inline_id(value)
    return text.upper() if text else ""


def _cli_name_key(value: object | None) -> str:
    return "".join(ch for ch in _clean_cli_name(value) if ch.isalnum())


def _cli_initial_alias(value: object | None) -> str:
    tokens = [token for token in re.split(r"[^A-Z0-9]+", _clean_cli_name(value)) if token]
    if len(tokens) < 2:
        return ""
    initials = " ".join(token[0] for token in tokens[:-1] if token)
    return f"{initials} {tokens[-1]}".strip()


def _cli_name_score(value: object | None) -> tuple[int, int, int]:
    tokens = [token for token in re.split(r"[^A-Z0-9]+", _clean_cli_name(value)) if token]
    long_tokens = sum(1 for token in tokens if len(token) > 1)
    return (long_tokens, len(tokens), len("".join(tokens)))


def _build_cli_name_maps(rows) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    alias_map = {key: _clean_cli_name(value) for key, value in CLI_NAME_MANUAL_ALIASES.items() if value}
    names_by_id: dict[str, set[str]] = {}
    for cli_name, cli_id in rows:
        name_text = _clean_cli_name(cli_name)
        id_text = _clean_cli_id(cli_id)
        if name_text:
            alias_map.setdefault(_cli_name_key(name_text), name_text)
        if id_text and name_text:
            names_by_id.setdefault(id_text, set()).add(name_text)

    canonical_by_id: dict[str, str] = {}
    id_sets_by_name_key: dict[str, set[str]] = {}
    for id_text, names in names_by_id.items():
        canonical = max(names, key=_cli_name_score)
        canonical_by_id[id_text] = canonical
        alias_map[_cli_name_key(canonical)] = canonical
        alias_name = _cli_initial_alias(canonical)
        if alias_name:
            alias_map.setdefault(_cli_name_key(alias_name), canonical)
        id_sets_by_name_key.setdefault(_cli_name_key(canonical), set()).add(id_text)

    id_by_name: dict[str, str] = {
        key: next(iter(ids))
        for key, ids in id_sets_by_name_key.items()
        if key and len(ids) == 1
    }
    return canonical_by_id, alias_map, id_by_name


def _canonicalize_cli_name(
    cli_name: object | None,
    cli_id: object | None = None,
    *,
    canonical_by_id: dict[str, str] | None = None,
    alias_map: dict[str, str] | None = None,
    id_by_name: dict[str, str] | None = None,
) -> tuple[str | None, str | None]:
    raw_name, inline_id = _split_cli_name_and_inline_id(cli_name)
    id_text = _clean_cli_id(cli_id) or _clean_cli_id(inline_id)
    name_text = _clean_cli_name(raw_name)
    lookup = dict(CLI_NAME_MANUAL_ALIASES)
    if alias_map:
        lookup.update(alias_map)

    if id_text and canonical_by_id and canonical_by_id.get(id_text):
        name_text = canonical_by_id[id_text]
    elif name_text:
        mapped = lookup.get(_cli_name_key(name_text))
        if mapped:
            name_text = _clean_cli_name(mapped)

    if not id_text and name_text and id_by_name:
        inferred_id = id_by_name.get(_cli_name_key(name_text))
        if inferred_id:
            id_text = inferred_id
            if canonical_by_id and canonical_by_id.get(id_text):
                name_text = canonical_by_id[id_text]

    return name_text or None, id_text or None


def format_cli_label(cli_name: str | None, cli_id: str | None = None) -> str:
    name_text, id_text = _canonicalize_cli_name(cli_name, cli_id)
    if name_text and id_text:
        return f"{name_text} ({id_text})"
    return name_text or id_text or ""


templates.env.filters["cli_label"] = format_cli_label


def _employee_cli_label(employee: Employee) -> str:
    cli_name, cli_id = _canonicalize_cli_name(employee.cli, employee.cli_id)
    return format_cli_label(cli_name, cli_id)


def _employee_cli_key(employee: Employee) -> str:
    cli_name, cli_id = _canonicalize_cli_name(employee.cli, employee.cli_id)
    return (cli_id or "").lower() or _cli_name_key(cli_name).lower()


def _cli_choice_rows(session: Session) -> list[dict[str, str]]:
    rows = session.exec(select(Employee.cli, Employee.cli_id).distinct()).all()
    canonical_by_id, alias_map, id_by_name = _build_cli_name_maps(rows)
    unique: dict[str, dict[str, str]] = {}
    for cli_name, cli_id in rows:
        name_text, id_text = _canonicalize_cli_name(
            cli_name,
            cli_id,
            canonical_by_id=canonical_by_id,
            alias_map=alias_map,
            id_by_name=id_by_name,
        )
        label = format_cli_label(name_text, id_text).strip()
        if not label:
            continue
        key = (id_text or "").lower() or _cli_name_key(name_text).lower()
        if not key or key in unique:
            continue
        unique[key] = {
            "label": label,
            "name": name_text or "",
            "cli_id": id_text or "",
        }
    return sorted(unique.values(), key=lambda item: item["label"].lower())


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


def _format_export_report_date_for_filename(value: object | None) -> str:
    text = _normalize_export_text(value)
    if not text:
        return ""
    if ":" in text:
        text = _normalize_export_text(text.split(":", 1)[1])
    for parser in (
        lambda raw: date.fromisoformat(raw),
        lambda raw: datetime.strptime(raw, "%d-%m-%Y").date(),
        lambda raw: datetime.strptime(raw, "%d/%m/%Y").date(),
    ):
        try:
            return parser(text).strftime("%d_%m_%Y")
        except ValueError:
            continue
    return ""


def _build_pdf_export_filename(title: object | None, report_date_label: object | None) -> str:
    safe_title = _sanitize_export_title(title)
    if safe_title == "SSTS PF Entering Speed Daily Report":
        report_date_suffix = _format_export_report_date_for_filename(report_date_label)
        if report_date_suffix:
            return f"{safe_title}_{report_date_suffix}.pdf"
    return _sanitize_export_filename(title, "pdf")


def _normalize_export_text(value: object | None) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _parse_export_report_date(value: object | None) -> str:
    text = _normalize_export_text(value)
    if not text:
        return ""
    if ":" in text:
        return text
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


def _coerce_export_table_payload(payload: object) -> tuple[str, list[str], list[list[str]], str, list[list[str]]]:
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Export payload must be an object.")

    headers_payload = payload.get("headers")
    rows_payload = payload.get("rows")
    cell_classes_payload = payload.get("cell_classes")
    if not isinstance(headers_payload, list) or not headers_payload:
        raise HTTPException(status_code=400, detail="Export requires at least one column.")
    if rows_payload is not None and not isinstance(rows_payload, list):
        raise HTTPException(status_code=400, detail="Export rows payload is invalid.")
    if cell_classes_payload is not None and not isinstance(cell_classes_payload, list):
        raise HTTPException(status_code=400, detail="Export cell classes payload is invalid.")

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
    cell_classes: list[list[str]] = []
    for raw_row in cell_classes_payload or []:
        if not isinstance(raw_row, list):
            continue
        normalized = [_normalize_export_text(cell) for cell in raw_row[:width]]
        if len(normalized) < width:
            normalized.extend([""] * (width - len(normalized)))
        cell_classes.append(normalized)
    while len(cell_classes) < len(rows):
        cell_classes.append([""] * width)
    return title, headers, rows, report_date_label, cell_classes


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
    cell_classes: list[list[str]] | None = None,
) -> bytes:
    page_width = 842.0
    page_height = 595.0
    margin_left = 26.0
    margin_right = 26.0
    footer_note = ""
    normalized_title = " ".join(str(title or "").split()).lower()
    if normalized_title.startswith("ssts pf "):
        footer_note = "* all data is taken from SSTS site based on data captured by the GPS tracking Device"
    margin_bottom = 38.0 if footer_note else 24.0
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
        cell_class_row: list[str] | None = None,
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

        for cell_index, (width, cell_lines) in enumerate(zip(widths, cells)):
            commands.append(f"{x + width:.2f} {bottom_y:.2f} m {x + width:.2f} {top_y:.2f} l S")
            text_x = x + padding_x
            text_y = top_y - padding_y - font_size
            current_text_color = text_color
            current_font_name = font_name
            current_font_size = font_size
            class_text = (cell_class_row[cell_index] if cell_class_row and cell_index < len(cell_class_row) else "").lower()
            if "pf-speed-alert" in class_text or "station-alert" in class_text:
                current_text_color = (0.769, 0.102, 0.102)
            if "crew-alert" in class_text:
                current_text_color = (0.027, 0.259, 0.522)
            if "pf-speed-alert" in class_text:
                current_font_name = "F2"
                current_font_size = font_size + 1.5
            if "station-alert" in class_text:
                current_font_name = "F2"
            if "crew-alert" in class_text:
                current_font_name = "F2"
            for line in cell_lines:
                add_text(commands, current_font_name, current_font_size, text_x, text_y, line, current_text_color)
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
            pdf_report_label = report_date_label
            if ":" not in pdf_report_label:
                pdf_report_label = f"Date: {pdf_report_label}"
            add_text(
                commands,
                "F1",
                9.4,
                margin_left,
                548.0,
                pdf_report_label,
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
        if footer_note:
            add_text(
                commands,
                "F1",
                8.6,
                margin_left,
                14.0,
                footer_note,
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
            source_row_index = sum(len(page) for page in pages[: page_index - 1]) + row_index
            current_y = draw_row(
                commands,
                row_layout,
                current_y,
                fill_color=(0.968, 0.980, 0.992) if row_index % 2 == 0 else (1.000, 1.000, 1.000),
                border_color=(0.792, 0.867, 0.925),
                text_color=(0.122, 0.180, 0.239),
                font_name="F1",
                cell_class_row=(cell_classes[source_row_index] if cell_classes and source_row_index < len(cell_classes) else None),
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
SSTS_API_POSITIONS_URL = f"{SSTS_API_BASE_URL}/timetable/tc/positions"
SSTS_API_CREW_URL = f"{SSTS_API_BASE_URL}/crew"
SSTS_API_USER = os.getenv("SSTS_API_USER", "srdeeopsdah@gmail.com")
SSTS_API_PASSWORD = os.getenv("SSTS_API_PASSWORD", "sdah1234")
SSTS_OFFLINE_THRESHOLD_MINUTES = 120
SSTS_RECENTLY_ONLINE_THRESHOLD_MINUTES = 5
SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES = 300
SSTS_RECENT_OFFLINE_MAX_MINUTES = 24 * 60
SSTS_REFRESH_INTERVAL_MINUTES = 5
SSTS_BACKGROUND_SYNC_INTERVAL_MINUTES = 60
SSTS_SNAPSHOT_RETENTION_DAYS = 7
SSTS_PF_REPORT_CACHE_TTL_MINUTES = 20
SSTS_PF_ANALYSIS_TASK_TTL_MINUTES = 180
SSTS_EXCLUDED_RAKE_NAMES = {"TEST1", "TEST2"}
SSTS_PF_SPIKE_FILTER_TRAIN_OVERRIDES: dict[str, set[str]] = {
    "2026-05-29": {
        "31233",
        "31527",
        "31528",
        "32216",
        "32248",
        "32412",
        "33231",
        "33320",
        "34540",
    },
    "2026-05-31": {
        "34856",
        "34919",
    },
    "2026-06-08": {
        "31527",
        "32214",
        "34526",
    }
}
IST = timezone(timedelta(hours=5, minutes=30))
_SSTS_PF_REPORT_CACHE: dict[str, tuple[datetime, dict[str, object]]] = {}
_SSTS_PF_ANALYSIS_TASKS: dict[str, dict[str, object]] = {}
_SSTS_PF_ANALYSIS_LOCK = threading.Lock()
_SSTS_CREW_CACHE: tuple[datetime, dict[str, str]] | None = None
_SSTS_BACKGROUND_SYNC_STOP = threading.Event()
_SSTS_BACKGROUND_SYNC_THREAD: threading.Thread | None = None
SSTS_PF_CREW_ID_OVERRIDES = {
    "KUNDAN KUMAR": "SDAH1898",
    "AMIT KUMAR": "SDAH2345",
}


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


def _top_performer_blank_state() -> dict[str, object]:
    return {
        "saved_at_label": "",
        "warnings": [],
        "minimum_runs": 3,
        "summary": None,
        "results": [],
        "comparison": None,
    }


def _top_performer_context(request: Request, **overrides: object) -> dict[str, object]:
    context = {
        "request": request,
        "active_page": "top_performer",
        **_top_performer_blank_state(),
    }
    context.update(overrides)
    return context


def _ensure_top_performer_dirs() -> None:
    TOP_PERFORMER_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOP_PERFORMER_PHOTO_DIR.mkdir(parents=True, exist_ok=True)


def _load_top_performer_store() -> dict[str, object]:
    if not TOP_PERFORMER_STORE_PATH.exists():
        return _top_performer_blank_state()
    try:
        payload = json.loads(TOP_PERFORMER_STORE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return _top_performer_blank_state()
    if not isinstance(payload, dict):
        return _top_performer_blank_state()
    store = _top_performer_blank_state()
    for key in store:
        if key in payload:
            store[key] = payload[key]
    return store


def _save_top_performer_store(payload: dict[str, object]) -> None:
    _ensure_top_performer_dirs()
    TOP_PERFORMER_STORE_PATH.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")


def _top_performer_slug(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", str(value or "").strip().lower()).strip("-") or "crew"


def _top_performer_display_number(value: object) -> object:
    if value in (None, ""):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value).strip()
    if math.isfinite(number) and number.is_integer():
        return int(number)
    return round(number, 2)


def _top_performer_numeric(value: object) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).strip().replace(",", "")
    if not cleaned:
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return 0.0
    try:
        return float(match.group(0))
    except ValueError:
        return 0.0


def _top_performer_normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _top_performer_read_tabular_file(upload: UploadFile) -> tuple[list[list[object]], str]:
    filename = upload.filename or "uploaded_file"
    content = upload.file.read()
    upload.file.seek(0)
    if filename.lower().endswith(".csv"):
        text = ""
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                text = content.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        rows = [list(row) for row in csv.reader(text.splitlines())]
        return rows, filename
    wb = load_workbook(filename=BytesIO(content), data_only=True)
    ws = wb.active
    return [list(row) for row in ws.iter_rows(values_only=True)], filename


def _top_performer_header_aliases() -> dict[str, tuple[str, ...]]:
    return {
        "crew_name": ("crewname", "crew", "name", "lpname", "loco", "motorman", "employee", "employeename"),
        "runs": ("runs", "run", "totalruns", "noofruns", "numberofruns", "trip", "trips", "totaltrip"),
        "total_score": ("totalscore", "score", "marks", "totalmarks", "grandtotal", "overallscore"),
        "bft": ("bft", "bftscore"),
        "bpt": ("bpt", "bptscore"),
        "speed": ("speed", "speedscore"),
        "platform": ("platform", "platformscore", "pf", "pfscore"),
        "emergency": ("emergency", "emergencyscore"),
        "cautious": ("cautious", "cautiousscore", "caution", "cautionscore"),
        "punctuality": ("punctuality", "punctualityscore", "punctual"),
        "report_date": ("reportdate", "date", "day"),
    }


def _top_performer_find_header(rows: list[list[object]]) -> tuple[int, dict[str, int]]:
    aliases = _top_performer_header_aliases()
    best_index = 0
    best_map: dict[str, int] = {}
    for row_index, row in enumerate(rows[:10]):
        header_map: dict[str, int] = {}
        for col_index, cell in enumerate(row):
            normalized = _top_performer_normalize_header(cell)
            if not normalized:
                continue
            for key, options in aliases.items():
                if normalized in options and key not in header_map:
                    header_map[key] = col_index
                    break
        if len(header_map) > len(best_map):
            best_index = row_index
            best_map = header_map
    return best_index, best_map


def _top_performer_guess_report_date(rows: list[list[object]], filename: str) -> str:
    text_candidates = [filename]
    for row in rows[:5]:
        for cell in row[:5]:
            if isinstance(cell, datetime):
                return cell.strftime("%d-%m-%Y")
            if isinstance(cell, date):
                return cell.strftime("%d-%m-%Y")
            if cell not in (None, ""):
                text_candidates.append(str(cell))
    for text in text_candidates:
        match = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{2,4})", text)
        if not match:
            continue
        try:
            return _excel_to_date(match.group(1)).strftime("%d-%m-%Y")  # type: ignore[union-attr]
        except Exception:
            continue
    return ""


def _top_performer_month_year_label(report_date: object | None, filename: str) -> str:
    parsed_date: date | None = None
    if report_date:
        try:
            parsed_date = datetime.strptime(str(report_date), "%d-%m-%Y").date()
        except ValueError:
            try:
                parsed_date = date.fromisoformat(str(report_date))
            except ValueError:
                parsed_date = None
    if parsed_date is None:
        date_texts: list[str] = []
        for match in re.finditer(r"\d{4}-\d{2}-\d{2}", filename):
            date_texts.append(match.group(0))
        for match in re.finditer(r"\d{2}[/-]\d{2}[/-]\d{2,4}", filename):
            date_texts.append(match.group(0))
        for text in date_texts:
            try:
                parsed = _excel_to_date(text)
            except Exception:
                continue
            if parsed:
                parsed_date = parsed
                break
    if parsed_date is None:
        return ""
    return parsed_date.strftime("%B %Y")


def _top_performer_card_heading(month_year_label: str) -> str:
    if not month_year_label:
        return "Crew Ranking Report"
    month_name, year = month_year_label.rsplit(" ", 1)
    return f"CREW RANKING REPORT MONTH OF {month_name}, {year}"


def _top_performer_poster_heading(prefix: str, month_year_label: str) -> str:
    if not month_year_label:
        return prefix
    return f"{prefix} - {month_year_label.upper()}"


def _top_performer_photo_lookup() -> dict[str, str]:
    if not TOP_PERFORMER_PHOTO_DIR.exists():
        return {}
    lookup: dict[str, str] = {}
    for path in TOP_PERFORMER_PHOTO_DIR.iterdir():
        if path.is_file():
            lookup[path.stem] = f"/static/top_performer_photos/{path.name}"
    return lookup


def _build_top_performer_result(rows: list[list[object]], filename: str, minimum_runs: int) -> tuple[dict[str, object], list[str]]:
    header_index, header_map = _top_performer_find_header(rows)
    warnings: list[str] = []
    if "crew_name" not in header_map:
        raise ValueError(f"Could not find a crew name column in {filename}.")
    if "total_score" not in header_map:
        raise ValueError(f"Could not find a total score column in {filename}.")
    photo_lookup = _top_performer_photo_lookup()
    parsed_rows: list[dict[str, object]] = []
    for row in rows[header_index + 1 :]:
        if not any(cell not in (None, "") for cell in row):
            continue

        def get_value(column: str) -> object:
            index = header_map.get(column)
            if index is None or index >= len(row):
                return ""
            return row[index]

        crew_name = str(get_value("crew_name") or "").strip()
        if not crew_name:
            continue
        runs_value = _top_performer_numeric(get_value("runs"))
        parsed_rows.append(
            {
                "crew_name": crew_name,
                "runs": int(runs_value) if runs_value.is_integer() else round(runs_value, 2),
                "total_score": _top_performer_display_number(get_value("total_score")),
                "total_score_value": _top_performer_numeric(get_value("total_score")),
                "bft": _top_performer_display_number(get_value("bft")),
                "bpt": _top_performer_display_number(get_value("bpt")),
                "speed": _top_performer_display_number(get_value("speed")),
                "platform": _top_performer_display_number(get_value("platform")),
                "emergency": _top_performer_display_number(get_value("emergency")),
                "cautious": _top_performer_display_number(get_value("cautious")),
                "punctuality": _top_performer_display_number(get_value("punctuality")),
                "photo_url": photo_lookup.get(_top_performer_slug(crew_name), ""),
            }
        )
    eligible_rows = [row for row in parsed_rows if _top_performer_numeric(row.get("runs")) >= minimum_runs]
    eligible_rows.sort(
        key=lambda item: (
            -_top_performer_numeric(item.get("total_score_value")),
            -_top_performer_numeric(item.get("runs")),
            str(item.get("crew_name") or "").lower(),
        )
    )
    top_rows: list[dict[str, object]] = []
    for index, row in enumerate(eligible_rows[:10], start=1):
        item = dict(row)
        item["rank"] = index
        item.pop("total_score_value", None)
        top_rows.append(item)
    if "runs" not in header_map:
        warnings.append(f'"{filename}" does not include a runs column, so all rows were treated as zero runs.')
    report_date_label = _top_performer_guess_report_date(rows, filename)
    month_year_label = _top_performer_month_year_label(report_date_label, filename)
    title = _top_performer_card_heading(month_year_label)
    return (
        {
            "filename": filename,
            "title": title,
            "poster_title": _top_performer_poster_heading("BEST PERFORMERS", month_year_label),
            "report_date": report_date_label,
            "row_count": len(parsed_rows),
            "eligible_count": len(eligible_rows),
            "top_rows": top_rows,
        },
        warnings,
    )


def _build_top_performer_comparison(
    previous_rows: list[list[object]],
    previous_filename: str,
    current_rows: list[list[object]],
    current_filename: str,
    minimum_runs: int,
) -> tuple[dict[str, object], list[str]]:
    previous_result, previous_warnings = _build_top_performer_result(previous_rows, previous_filename, minimum_runs)
    current_result, current_warnings = _build_top_performer_result(current_rows, current_filename, minimum_runs)
    header_index_prev, header_map_prev = _top_performer_find_header(previous_rows)
    header_index_curr, header_map_curr = _top_performer_find_header(current_rows)

    def parse_all(rows_data: list[list[object]], header_index: int, header_map: dict[str, int]) -> dict[str, dict[str, object]]:
        photo_lookup = _top_performer_photo_lookup()
        records: dict[str, dict[str, object]] = {}
        for row in rows_data[header_index + 1 :]:
            if not any(cell not in (None, "") for cell in row):
                continue
            name_index = header_map.get("crew_name")
            if name_index is None or name_index >= len(row):
                continue
            crew_name = str(row[name_index] or "").strip()
            if not crew_name:
                continue
            runs_value = _top_performer_numeric(row[header_map["runs"]]) if "runs" in header_map and header_map["runs"] < len(row) else 0.0
            score_value = _top_performer_numeric(row[header_map["total_score"]]) if header_map["total_score"] < len(row) else 0.0
            if runs_value < minimum_runs:
                continue
            records[_top_performer_slug(crew_name)] = {
                "crew_name": crew_name,
                "runs": int(runs_value) if runs_value.is_integer() else round(runs_value, 2),
                "score": round(score_value, 2) if not score_value.is_integer() else int(score_value),
                "score_value": score_value,
                "photo_url": photo_lookup.get(_top_performer_slug(crew_name), ""),
            }
        return records

    previous_all = parse_all(previous_rows, header_index_prev, header_map_prev)
    current_all = parse_all(current_rows, header_index_curr, header_map_curr)
    comparison_rows: list[dict[str, object]] = []
    improved_rows: list[dict[str, object]] = []
    for slug, current_row in current_all.items():
        previous_row = previous_all.get(slug)
        if not previous_row:
            continue
        change_value = _top_performer_numeric(current_row.get("score_value")) - _top_performer_numeric(previous_row.get("score_value"))
        if change_value > 0:
            status = "Improved"
        elif change_value < 0:
            status = "Declined"
        else:
            status = "No Change"
        row_payload = {
            "crew_name": current_row.get("crew_name"),
            "photo_url": current_row.get("photo_url"),
            "previous_runs": previous_row.get("runs"),
            "current_runs": current_row.get("runs"),
            "previous_score": previous_row.get("score"),
            "current_score": current_row.get("score"),
            "score_change": _top_performer_display_number(change_value),
            "score_change_value": change_value,
            "status": status,
        }
        comparison_rows.append(row_payload)
        if change_value > 0:
            improved_rows.append(dict(row_payload))
    comparison_rows.sort(
        key=lambda item: (
            -_top_performer_numeric(item.get("score_change_value")),
            -_top_performer_numeric(item.get("current_score")),
            str(item.get("crew_name") or "").lower(),
        )
    )
    improved_rows.sort(
        key=lambda item: (
            -_top_performer_numeric(item.get("score_change_value")),
            -_top_performer_numeric(item.get("current_score")),
            str(item.get("crew_name") or "").lower(),
        )
    )
    display_rows = improved_rows if improved_rows else comparison_rows
    top_display_rows = display_rows[:10]
    for index, row in enumerate(top_display_rows, start=1):
        row["rank"] = index
        row.pop("score_change_value", None)
    comparison_month_year = _top_performer_month_year_label(
        current_result.get("report_date"),
        current_filename,
    )
    return (
        {
            "poster_title": _top_performer_poster_heading("TOP TEN IMPROVED CREW", comparison_month_year),
            "previous_filename": previous_filename,
            "current_filename": current_filename,
            "previous_report_date": previous_result.get("report_date", ""),
            "current_report_date": current_result.get("report_date", ""),
            "previous_eligible_count": len(previous_all),
            "current_eligible_count": len(current_all),
            "matched_count": len(comparison_rows),
            "improved_count": len(improved_rows),
            "rows": top_display_rows,
        },
        previous_warnings + current_warnings,
    )

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


def build_cli_distribution(
    employees: list[Employee],
    cli_bio_reference_rows: list[dict[str, str]] | None = None,
) -> list[dict[str, int | str]]:
    """Aggregate gradation counts per CLI and preserve uploaded CLI master IDs."""
    reference_rows = cli_bio_reference_rows or []
    reference_lookup: dict[str, tuple[str, str]] = {}
    for row in reference_rows:
        ref_name, ref_id = _canonicalize_cli_name(row.get("cli_name"), row.get("cli_id"))
        ref_key = (ref_id or "").lower() or _cli_name_key(ref_name).lower()
        if not ref_key:
            continue
        reference_lookup[ref_key] = (ref_name or "", ref_id or "")

    dist: dict[str, dict[str, int | str]] = {}
    for e in employees:
        cli_name, cli_id = _canonicalize_cli_name(e.cli, e.cli_id)
        cli_key = (cli_id or "").lower() or _cli_name_key(cli_name).lower() or "unassigned"
        ref_name, ref_id = reference_lookup.get(cli_key, ("", ""))
        label = ref_name or cli_name or "Unassigned"
        grad = (e.gradation or "").strip().upper()
        grad_key = grad[0] if grad else ""
        if cli_key not in dist:
            dist[cli_key] = {
                "cli": label,
                "cli_id": ref_id or cli_id or "",
                "A": 0,
                "B": 0,
                "C": 0,
                "total": 0,
                "total_staff": 0,
            }
        if not dist[cli_key]["cli"] and cli_name:
            dist[cli_key]["cli"] = cli_name
        if not dist[cli_key]["cli_id"] and cli_id:
            dist[cli_key]["cli_id"] = cli_id
        if grad_key in ("A", "B", "C"):
            dist[cli_key][grad_key] += 1  # type: ignore[index]
            dist[cli_key]["total"] += 1  # type: ignore[index]
        dist[cli_key]["total_staff"] += 1  # type: ignore[index]

    for ref_name, ref_id in reference_lookup.values():
        cli_key = (ref_id or "").lower() or _cli_name_key(ref_name).lower()
        if cli_key not in dist:
            dist[cli_key] = {
                "cli": ref_name or "Unassigned",
                "cli_id": ref_id or "",
                "A": 0,
                "B": 0,
                "C": 0,
                "total": 0,
                "total_staff": 0,
            }

    # Collapse rows that share the same canonical CLI name when some rows are missing CLI ID.
    merged_dist: dict[str, dict[str, int | str]] = {}
    for _, counts in dist.items():
        cli_name = str(counts.get("cli") or "")
        cli_id = str(counts.get("cli_id") or "")
        merge_key = cli_id.lower() or _cli_name_key(cli_name).lower() or "unassigned"
        existing = merged_dist.get(merge_key)
        if existing is None:
            merged_dist[merge_key] = dict(counts)
            continue
        if not existing.get("cli") and cli_name:
            existing["cli"] = cli_name
        if not existing.get("cli_id") and cli_id:
            existing["cli_id"] = cli_id
        for bucket in ("A", "B", "C", "total", "total_staff"):
            existing[bucket] = int(existing.get(bucket) or 0) + int(counts.get(bucket) or 0)

    # Also collapse rows by canonical CLI name so a zero-count reference row with CLI ID
    # does not stay separate from a populated row that only has the same CLI name.
    final_dist: dict[str, dict[str, int | str]] = {}
    for _, counts in merged_dist.items():
        cli_name = str(counts.get("cli") or "")
        cli_id = str(counts.get("cli_id") or "")
        merge_key = _cli_name_key(cli_name).lower() or cli_id.lower() or "unassigned"
        existing = final_dist.get(merge_key)
        if existing is None:
            final_dist[merge_key] = dict(counts)
            continue
        if not existing.get("cli") and cli_name:
            existing["cli"] = cli_name
        if not existing.get("cli_id") and cli_id:
            existing["cli_id"] = cli_id
        for bucket in ("A", "B", "C", "total", "total_staff"):
            existing[bucket] = int(existing.get(bucket) or 0) + int(counts.get(bucket) or 0)

    return [
        {
            "cli": counts["cli"],
            "cli_id": counts["cli_id"],
            "A": counts["A"],
            "B": counts["B"],
            "C": counts["C"],
            "total": counts["total"],
            "total_staff": counts["total_staff"],
        }
        for _, counts in sorted(
            final_dist.items(),
            key=lambda item: (
                str(item[1].get("cli") or "").lower(),
                str(item[1].get("cli_id") or "").lower(),
            ),
        )
    ]

def build_working_location_summary(
    employees: list[Employee],
) -> tuple[list[str], list[dict[str, object]]]:
    """Aggregate CCR staff counts by working location and role."""
    dynamic_roles = sorted({e.role for e in employees if e.role not in ROLE_ORDER})
    role_headers = ROLE_ORDER + [r for r in dynamic_roles if r not in ROLE_ORDER]
    working_summary: list[dict[str, object]] = []
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
            continue
        loc = allowed_norm[loc_key]
        working_map.setdefault(loc, {}).setdefault(e.role, 0)
        working_map[loc][e.role] += 1
    for loc in sorted(working_map.keys(), key=lambda value: value.lower()):
        counts = {role: working_map[loc].get(role, 0) for role in role_headers}
        working_summary.append(
            {
                "working_at": loc,
                "counts": counts,
                "total": sum(counts.values()),
            }
        )
    return role_headers, working_summary


def _sync_retired_employees(session: Session, as_of: date | None = None) -> int:
    return purge_retired_employees(session, as_of or date.today())


def _sync_employee_duplicates(session: Session) -> int:
    plan, _, _ = _build_combined_cleanup_view(session)
    mergeable_plan = [
        item
        for item in plan
        if isinstance(item.get("keep"), dict) and item.get("keep") and list(item.get("remove") or [])
    ]
    if not mergeable_plan:
        return 0
    details: list[str] = []
    return _apply_duplicate_cleanup_plan(session, mergeable_plan, details)


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


def _ssts_offline_transition_time(lastupdate: datetime | None) -> datetime | None:
    lastupdate_utc = _ensure_utc(lastupdate)
    if lastupdate_utc is None:
        return None
    return lastupdate_utc + timedelta(minutes=SSTS_OFFLINE_THRESHOLD_MINUTES)


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


def _ssts_request_with_retry(req: urlrequest.Request, *, expects_json: bool = True) -> object:
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            with urlrequest.urlopen(req, timeout=30) as response:
                body = response.read().decode("utf-8", "replace")
            return json.loads(body) if expects_json else body
        except (urlerror.URLError, TimeoutError, ConnectionResetError, OSError, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt >= 2:
                break
            time.sleep(1.2 * (attempt + 1))
    if last_error is not None:
        raise RuntimeError(f"SSTS request failed after retries: {last_error}") from last_error
    raise RuntimeError("SSTS request failed after retries.")


def _ssts_post_json(url: str, payload: dict[str, object], headers: dict[str, str] | None = None) -> dict[str, object]:
    body = json.dumps(payload).encode("utf-8")
    request_headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    if headers:
        request_headers.update(headers)
    req = urlrequest.Request(url, data=body, headers=request_headers)
    response = _ssts_request_with_retry(req)
    if not isinstance(response, dict):
        raise RuntimeError("Unexpected SSTS POST response format.")
    return response


def _ssts_get_json(url: str, headers: dict[str, str] | None = None) -> object:
    request_headers = {"User-Agent": "Mozilla/5.0"}
    if headers:
        request_headers.update(headers)
    req = urlrequest.Request(url, headers=request_headers)
    return _ssts_request_with_retry(req)


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


def _coerce_int(value: object) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _pf_chart_speed_kmph(point: dict[str, object]) -> float | None:
    try:
        raw_speed = float(point.get("speed") or 0)
    except (TypeError, ValueError):
        return None
    device_id = _coerce_int(point.get("deviceid"))
    if device_id not in {3, 10, 11, 13}:
        raw_speed *= 1.852
    return raw_speed


def _pf_chart_distance_km(point: dict[str, object]) -> float | None:
    for key in ("totalDistance", "totaldistance", "distance", "dist", "km", "distance_km"):
        try:
            value = point.get(key)
        except AttributeError:
            value = None
        if value in (None, ""):
            continue
        try:
            return float(value)
        except (TypeError, ValueError):
            continue
    raw_attributes = point.get("attributes")
    if raw_attributes not in (None, ""):
        try:
            attributes = json.loads(str(raw_attributes))
        except (TypeError, ValueError, json.JSONDecodeError):
            attributes = {}
        for key in ("totalDistance", "totaldistance", "distance"):
            value = attributes.get(key)
            if value in (None, ""):
                continue
            try:
                return float(value)
            except (TypeError, ValueError):
                continue
    return None


def _pf_chart_distance_series_km(chart_points: list[dict[str, object]]) -> list[float | None]:
    raw_values = [_pf_chart_distance_km(point) for point in chart_points]
    first_non_null = next((value for value in raw_values if value is not None), None)
    if first_non_null is None:
        return [None for _ in chart_points]

    if first_non_null > 1000:
        baseline = first_non_null
        return [
            None if value is None else max(0.0, (value - baseline) / 1000.0)
            for value in raw_values
        ]

    return [
        None if value is None else float(value) / 1000.0
        for value in raw_values
    ]


def _pf_chart_time_label(point: dict[str, object]) -> str:
    for key in ("gpstime", "gps_time", "time", "device_time", "servertime", "updatedon", "updated_at"):
        value = point.get(key)
        text = str(value or "").strip()
        if not text:
            continue
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.astimezone(IST)
            return parsed.strftime("%H:%M:%S")
        except ValueError:
            parts = text.split()
            if len(parts) >= 2:
                return parts[-1][:8]
            return text[:8]
    return ""


def _pf_build_chart_link(row: dict[str, object]) -> str:
    query = urlparse.urlencode(
        {
            "train_date": str(row.get("train_date_iso") or row.get("train_date") or ""),
            "train_no": str(row.get("train_no") or ""),
            "device_id": str(row.get("device_id") or ""),
            "org": str(row.get("org") or ""),
            "dep": str(row.get("train_dep_raw") or row.get("dep") or ""),
            "dest": str(row.get("dest") or ""),
            "arr": str(row.get("train_arr_raw") or row.get("arr") or ""),
            "station": str(row.get("station") or ""),
            "start_pos": str(row.get("start_pos") or ""),
            "end_pos": str(row.get("end_pos") or ""),
        }
    )
    return f"/ssts-report/pf-chart?{query}"


def _pf_attach_chart_links(rows: object) -> object:
    if not isinstance(rows, list):
        return rows
    attached_rows: list[dict[str, object] | object] = []
    for row in rows:
        if not isinstance(row, dict):
            attached_rows.append(row)
            continue
        row_copy = dict(row)
        row_copy["chart_link"] = _pf_build_chart_link(row_copy)
        attached_rows.append(row_copy)
    return attached_rows


def _pf_attach_chart_links_by_train(rows_by_train: object) -> object:
    if not isinstance(rows_by_train, dict):
        return rows_by_train
    attached: dict[object, object] = {}
    for train_no, rows in rows_by_train.items():
        attached[train_no] = _pf_attach_chart_links(rows)
    return attached


def _pf_hydrate_chart_links_in_result(result: dict[str, object]) -> dict[str, object]:
    hydrated = dict(result)
    for key in (
        "pf_report_rows",
        "pf_daily_report_rows",
        "pf_detailed_daily_report_rows",
        "pf_detailed_daily_spike_rows",
        "pf_analysis_summary_rows",
    ):
        if key in hydrated:
            hydrated[key] = _pf_attach_chart_links(hydrated.get(key))
    for key in ("pf_analysis_detail_rows_by_train", "pf_detailed_detail_rows_by_train"):
        if key in hydrated:
            hydrated[key] = _pf_attach_chart_links_by_train(hydrated.get(key))
    return hydrated


def _pf_build_station_plot_bands(
    rows: list[dict[str, object]],
    selected_station: str,
    selected_start: int | None,
    selected_end: int | None,
) -> list[dict[str, object]]:
    plot_bands: list[dict[str, object]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        band_start = _coerce_int(row.get("start_pos"))
        band_end = _coerce_int(row.get("end_pos"))
        if band_start is None or band_end is None:
            continue
        if band_end < band_start:
            band_start, band_end = band_end, band_start
        station_name = str(row.get("station") or "").strip() or "STN"
        is_selected = (
            station_name == selected_station
            and selected_start is not None
            and selected_end is not None
            and band_start == min(selected_start, selected_end)
            and band_end == max(selected_start, selected_end)
        )
        arr_text = str(row.get("act_arr") or row.get("sch_arr") or "").strip()
        dep_text = str(row.get("act_dep") or row.get("sch_dep") or "").strip()
        label_lines = [station_name]
        if arr_text:
            label_lines.append(arr_text)
        if dep_text and dep_text != arr_text:
            label_lines.append(dep_text)
        plot_bands.append(
            {
                "from": band_start,
                "to": band_end,
                "isSelected": is_selected,
                "color": "rgba(134, 239, 172, 0.28)" if not is_selected else "rgba(253, 224, 71, 0.32)",
                "borderColor": "rgba(34, 197, 94, 0.38)" if not is_selected else "rgba(217, 119, 6, 0.60)",
                "borderWidth": 1,
                "label": {
                    "text": "<br/>".join(label_lines),
                    "useHTML": True,
                    "style": {
                        "color": "#14532d" if not is_selected else "#92400e",
                        "fontWeight": "700",
                        "fontSize": "11px",
                        "textAlign": "center",
                    },
                },
            }
        )
    return plot_bands


def _pf_chart_point_seconds(point: dict[str, object]) -> int | None:
    for key in ("gpstime", "gps_time", "time", "device_time", "servertime", "updatedon", "updated_at"):
        value = point.get(key)
        text = str(value or "").strip()
        if not text:
            continue
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.astimezone(IST)
            return (parsed.hour * 3600) + (parsed.minute * 60) + parsed.second
        except ValueError:
            parts = text.split()
            hhmmss = parts[-1][:8] if parts else text[:8]
            parsed_seconds = _parse_hms_seconds(hhmmss)
            if parsed_seconds is not None:
                return parsed_seconds
    return None


def _pf_monotonic_seconds(values: list[int | None]) -> list[int | None]:
    normalized: list[int | None] = []
    day_offset = 0
    previous_value: int | None = None
    for value in values:
        if value is None:
            normalized.append(None)
            continue
        adjusted_value = value + day_offset
        if previous_value is not None and adjusted_value < previous_value - 43200:
            day_offset += 86400
            adjusted_value = value + day_offset
        normalized.append(adjusted_value)
        previous_value = adjusted_value
    return normalized


def _pf_fit_seconds_to_chart_window(value: int | None, chart_min: int, chart_max: int) -> int | None:
    if value is None:
        return None
    candidates = [value + (86400 * offset) for offset in (-1, 0, 1, 2)]
    return min(
        candidates,
        key=lambda candidate: (
            0 if chart_min <= candidate <= chart_max else min(abs(candidate - chart_min), abs(candidate - chart_max)),
            abs(candidate - chart_min),
        ),
    )


def _pf_find_nearest_chart_index(
    chart_seconds: list[int | None],
    target_seconds: int | None,
) -> int | None:
    if target_seconds is None:
        return None
    nearest_index: int | None = None
    nearest_gap: int | None = None
    for index, point_seconds in enumerate(chart_seconds):
        if point_seconds is None:
            continue
        gap = abs(point_seconds - target_seconds)
        if nearest_gap is None or gap < nearest_gap:
            nearest_index = index
            nearest_gap = gap
    return nearest_index


def _pf_interpolate_position(
    anchors: list[tuple[int, int]],
    original_position: int,
    chart_point_count: int,
) -> int:
    if not anchors:
        return max(0, min(chart_point_count - 1, original_position))
    if len(anchors) == 1:
        original_anchor, mapped_anchor = anchors[0]
        return max(0, min(chart_point_count - 1, mapped_anchor + (original_position - original_anchor)))
    if original_position <= anchors[0][0]:
        left_original, left_mapped = anchors[0]
        right_original, right_mapped = anchors[1]
    elif original_position >= anchors[-1][0]:
        left_original, left_mapped = anchors[-2]
        right_original, right_mapped = anchors[-1]
    else:
        left_original = left_mapped = right_original = right_mapped = 0
        for idx in range(1, len(anchors)):
            previous_original, previous_mapped = anchors[idx - 1]
            current_original, current_mapped = anchors[idx]
            if previous_original <= original_position <= current_original:
                left_original, left_mapped = previous_original, previous_mapped
                right_original, right_mapped = current_original, current_mapped
                break
    span_original = right_original - left_original
    if span_original == 0:
        mapped_position = left_mapped
    else:
        ratio = (original_position - left_original) / span_original
        mapped_position = left_mapped + ((right_mapped - left_mapped) * ratio)
    return max(0, min(chart_point_count - 1, int(round(mapped_position))))


def _pf_snap_station_window_to_stop(
    chart_points: list[dict[str, object]],
    start_index: int | None,
    end_index: int | None,
    stop_time_seconds: int | None,
) -> tuple[int | None, int | None]:
    if not chart_points or start_index is None or end_index is None:
        return start_index, end_index
    left = min(start_index, end_index)
    right = max(start_index, end_index)
    if left < 0 or right < 0:
        return start_index, end_index

    search_left = max(0, left - 18)
    search_right = min(len(chart_points) - 1, right + 24)
    midpoint = (left + right) / 2
    candidates: list[tuple[float, float, int]] = []
    for idx in range(search_left, search_right + 1):
        speed = _pf_chart_speed_kmph(chart_points[idx])
        if speed is None:
            continue
        if speed <= 12:
            candidates.append((speed, abs(idx - midpoint), idx))
    if not candidates:
        return start_index, end_index

    _, _, valley_index = min(candidates, key=lambda item: (item[0], item[1], item[2]))

    snapped_left = valley_index
    snapped_right = valley_index
    while snapped_left > search_left:
        speed = _pf_chart_speed_kmph(chart_points[snapped_left - 1])
        if speed is None or speed > 15:
            break
        snapped_left -= 1
    while snapped_right < search_right:
        speed = _pf_chart_speed_kmph(chart_points[snapped_right + 1])
        if speed is None or speed > 15:
            break
        snapped_right += 1

    if stop_time_seconds is not None and stop_time_seconds > 0:
        minimum_span = max(2, min(18, int(round(stop_time_seconds / 15))))
        current_span = snapped_right - snapped_left
        if current_span < minimum_span:
            pad = int(math.ceil((minimum_span - current_span) / 2))
            snapped_left = max(search_left, snapped_left - pad)
            snapped_right = min(search_right, snapped_right + pad)

    if abs(((snapped_left + snapped_right) / 2) - midpoint) >= 3:
        return snapped_left, snapped_right
    return start_index, end_index


def _pf_normalize_station_windows_to_chart(
    rows: list[dict[str, object]],
    chart_points: list[dict[str, object]],
    chart_point_count: int,
    selected_station: str,
    selected_start: int | None,
    selected_end: int | None,
) -> tuple[list[dict[str, object]], int | None, int | None]:
    if chart_point_count <= 1:
        return rows, selected_start, selected_end

    max_end = max(
        (
            max(
                _coerce_int(row.get("start_pos")) or 0,
                _coerce_int(row.get("end_pos")) or 0,
            )
            for row in rows
            if isinstance(row, dict)
        ),
        default=0,
    )
    if max_end <= 0 or max_end <= (chart_point_count - 1):
        return rows, selected_start, selected_end

    scale = (chart_point_count - 1) / max_end
    normalized_rows: list[dict[str, object]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        row_copy = dict(row)
        start_pos = _coerce_int(row.get("start_pos"))
        end_pos = _coerce_int(row.get("end_pos"))
        if start_pos is not None:
            row_copy["start_pos"] = int(round(start_pos * scale))
        if end_pos is not None:
            row_copy["end_pos"] = int(round(end_pos * scale))
        normalized_rows.append(row_copy)

    normalized_selected_start = int(round(selected_start * scale)) if selected_start is not None else None
    normalized_selected_end = int(round(selected_end * scale)) if selected_end is not None else None
    return normalized_rows, normalized_selected_start, normalized_selected_end


def _build_pf_positions_params(source: dict[str, object]) -> dict[str, object]:
    return {
        "train_date": source.get("train_date_iso") or source.get("train_date"),
        "train_no": source.get("train_no"),
        "device_id": source.get("device_id"),
        "org": source.get("org"),
        "dep": source.get("train_dep_raw"),
        "dest": source.get("dest"),
        "arr": source.get("train_arr_raw"),
    }


def _fetch_ssts_positions(source: dict[str, object], token: str) -> list[dict[str, object]]:
    response = _ssts_get_json_with_params(
        SSTS_API_POSITIONS_URL,
        _build_pf_positions_params(source),
        headers={"Authorization": token},
    )
    if not isinstance(response, list):
        return []
    return [point for point in response if isinstance(point, dict)]


def _normalize_pf_crew_name(value: object) -> str:
    text = str(value or "").strip().upper()
    if not text:
        return ""
    return re.sub(r"\s+", " ", text)


def _resolve_pf_crew_id(crew_name: object) -> str:
    normalized_name = _normalize_pf_crew_name(crew_name)
    if not normalized_name:
        return ""
    return SSTS_PF_CREW_ID_OVERRIDES.get(normalized_name, "")


def _build_pf_report_rows_for_train(
    train: dict[str, object],
    report_day: date,
    token: str,
) -> list[dict[str, object]]:
    base_row = {
        "report_date": report_day.strftime("%d-%m-%Y"),
        "train_date_iso": report_day.isoformat(),
        "train_no": str(train.get("train_no") or ""),
        "rake_no": str(train.get("device_name") or ""),
        "device_id": train.get("device_id"),
        "org": str(train.get("org") or ""),
        "dest": str(train.get("dest") or ""),
        "crew_name": str(train.get("crew_name") or ""),
        "train_dep_raw": train.get("dep"),
        "train_arr_raw": train.get("arr"),
    }
    base_row["chart_link"] = _pf_build_chart_link(base_row)
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
                "start_pos": "",
                "end_pos": "",
                "speed_at_600m": "",
                "speed_at_400m": "",
                "speed_at_265m": "",
                "speed_at_100m": "",
                "status_message": "Data not found or Device might be Offline",
            }
        ]
    detail_rows: list[dict[str, object]] = []
    for item in response:
        if not isinstance(item, dict):
            continue
        crew_name = str(item.get("crew_name") or "").strip()
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
                "start_pos": item.get("start_pos") if item.get("start_pos") is not None else "",
                "end_pos": item.get("end_pos") if item.get("end_pos") is not None else "",
                "speed_at_600m": item.get("speed_at_600m") if item.get("speed_at_600m") is not None else "",
                "speed_at_400m": item.get("speed_at_400m") if item.get("speed_at_400m") is not None else "",
                "speed_at_265m": item.get("speed_at_265m") if item.get("speed_at_265m") is not None else "",
                "speed_at_100m": item.get("speed_at_100m") if item.get("speed_at_100m") is not None else "",
                "crew_name": crew_name,
                "crew_id": _resolve_pf_crew_id(crew_name),
                "remarks": str(item.get("remarks") or ""),
                "status_message": "",
            }
        )
        detail_rows[-1]["chart_link"] = _pf_build_chart_link(detail_rows[-1])
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
            "start_pos": "",
            "end_pos": "",
            "speed_at_600m": "",
            "speed_at_400m": "",
            "speed_at_265m": "",
            "speed_at_100m": "",
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
    crew_lookup: dict[str, str] = {}
    try:
        crew_lookup = fetch_ssts_crew_lookup(token)
    except (urlerror.URLError, RuntimeError, ValueError, json.JSONDecodeError):
        crew_lookup = {}
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
    for row in rows:
        if not isinstance(row, dict):
            continue
        crew_name_key = _normalize_ssts_crew_name(row.get("crew_name"))
        row["crew_id"] = _resolve_pf_crew_id(row.get("crew_name")) or (crew_lookup.get(crew_name_key, "") if crew_name_key else "")
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


def _parse_pf_speed_threshold(value: object | None) -> int:
    try:
        threshold = int(str(value or "40").strip())
    except (TypeError, ValueError):
        threshold = 40
    if 40 <= threshold <= 50:
        return threshold
    if threshold > 50:
        return 51
    return 40


def _pf_speed_threshold_label(threshold: int) -> str:
    return "Above 50" if threshold > 50 else f"{threshold} and above"


def _pf_speed_matches_threshold(speed: float | None, threshold: int) -> bool:
    if speed is None:
        return False
    if threshold > 50:
        return speed > 50
    return speed >= threshold


def _pf_reference_speed(row: dict[str, object] | None) -> float | None:
    if not isinstance(row, dict):
        return None
    pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
    if pf_speed is not None:
        return pf_speed
    return _pf_speed_value(row.get("geofence_enter_speed"))


def _parse_hms_seconds(value: object | None) -> int | None:
    text = str(value or "").strip()
    if not text:
        return None
    parts = text.split(":")
    if len(parts) != 3:
        return None
    try:
        hours, minutes, seconds = (int(part) for part in parts)
    except ValueError:
        return None
    if hours < 0 or minutes < 0 or seconds < 0:
        return None
    return (hours * 3600) + (minutes * 60) + seconds


def _pf_suspected_spike_reason(
    row: dict[str, object],
    previous_row: dict[str, object] | None,
    next_row: dict[str, object] | None,
    threshold: int,
    chart_points: list[dict[str, object]] | None = None,
) -> str | None:
    pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
    if not _pf_speed_matches_threshold(pf_speed, threshold):
        return None

    assert pf_speed is not None
    geofence_speed = _pf_speed_value(row.get("geofence_enter_speed"))
    stop_time_seconds = _parse_hms_seconds(row.get("stop_time"))
    max_entry_speed = max(speed for speed in (pf_speed, geofence_speed) if speed is not None)

    start_pos = _coerce_int(row.get("start_pos"))
    if chart_points and start_pos is not None and start_pos >= 0:
        end_pos = _coerce_int(row.get("end_pos"))
        if (
            stop_time_seconds is not None
            and stop_time_seconds <= 90
            and len(chart_points) >= 8
            and (
                start_pos >= len(chart_points)
                or (end_pos is not None and end_pos >= len(chart_points))
            )
        ):
            pf_distance = _pf_speed_value(row.get("pf_distance"))
            trailing_window = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[max(0, len(chart_points) - 18) :]
            ]
            trailing_window = [speed for speed in trailing_window if speed is not None]
            if trailing_window and pf_distance is not None and 240 <= pf_distance <= 320:
                trailing_peak = max(trailing_window)
                trailing_end = trailing_window[-1]
                trailing_zero_run = 0
                trailing_longest_zero_run = 0
                sharp_rise_count = sum(
                    1
                    for idx in range(1, len(trailing_window))
                    if (trailing_window[idx] - trailing_window[idx - 1]) >= 8
                )
                sharp_drop_count = sum(
                    1
                    for idx in range(1, len(trailing_window))
                    if (trailing_window[idx] - trailing_window[idx - 1]) <= -6
                )
                for speed in trailing_window:
                    if speed <= 5:
                        trailing_zero_run += 1
                        trailing_longest_zero_run = max(trailing_longest_zero_run, trailing_zero_run)
                    else:
                        trailing_zero_run = 0
                if (
                    pf_speed >= max(40.0, threshold)
                    and trailing_peak >= pf_speed + 8
                    and (trailing_peak - trailing_end) >= 10
                    and sharp_rise_count >= 1
                    and sharp_drop_count >= 1
                ):
                    return "Station chart ended before the stop window and the tail showed a sharp spike/drop."
        window_end = min(start_pos, len(chart_points) - 1)
        window_start = max(0, window_end - 40)
        spike_window_end = min(
            len(chart_points) - 1,
            max(window_end, end_pos if end_pos is not None and end_pos >= 0 else window_end),
        )
        spike_window_start = max(1, window_end - 6)
        spike_window_limit = min(
            len(chart_points) - 1,
            max(
                spike_window_end,
                min(len(chart_points) - 1, window_end + 24),
            ),
        )
        for idx in range(spike_window_start, spike_window_limit):
            prev_speed = _pf_chart_speed_kmph(chart_points[idx - 1])
            peak_speed = _pf_chart_speed_kmph(chart_points[idx])
            next_speed = _pf_chart_speed_kmph(chart_points[idx + 1]) if idx + 1 < len(chart_points) else None
            next2_speed = _pf_chart_speed_kmph(chart_points[idx + 2]) if idx + 2 < len(chart_points) else None
            next3_speed = _pf_chart_speed_kmph(chart_points[idx + 3]) if idx + 3 < len(chart_points) else None
            if None in (prev_speed, peak_speed, next_speed):
                continue
            if (
                peak_speed >= max(45.0, threshold)
                and peak_speed - prev_speed >= 8
                and (
                    peak_speed - next_speed >= 8
                    or (
                        peak_speed - prev_speed >= 12
                        and (
                            (next2_speed is not None and next2_speed <= peak_speed - 15)
                            or (next3_speed is not None and next3_speed <= peak_speed - 25)
                        )
                    )
                )
                and (
                    next_speed <= peak_speed - 10
                    or (next2_speed is not None and next2_speed <= peak_speed - 15)
                    or (next3_speed is not None and next3_speed <= peak_speed - 25)
                )
            ):
                return "Sharp chart peak collapsed immediately before station entry."
            if (
                pf_speed <= 50
                and geofence_speed is not None
                and abs(pf_speed - geofence_speed) <= 6
                and peak_speed >= max(pf_speed, geofence_speed) + 10
                and peak_speed - prev_speed >= 8
                and next_speed <= peak_speed - 4
                and (
                    (next2_speed is not None and next2_speed <= peak_speed - 10)
                    or (next3_speed is not None and next3_speed <= peak_speed - 12)
                )
            ):
                return "Moderate-speed chart showed a short-lived local spike before settling back."

        entry_window_start = max(0, window_end - 12)
        entry_window_end = min(len(chart_points) - 1, spike_window_end + 3)
        entry_window_speeds = [
            _pf_chart_speed_kmph(point)
            for point in chart_points[entry_window_start : entry_window_end + 1]
        ]
        entry_window_speeds = [speed for speed in entry_window_speeds if speed is not None]
        if entry_window_speeds:
            entry_window_peak = max(entry_window_speeds)
            if (
                geofence_speed is not None
                and pf_speed >= max(45.0, threshold)
                and (pf_speed - geofence_speed) >= 18
                and entry_window_peak <= geofence_speed + 5
                and entry_window_peak <= pf_speed - 15
            ):
                return "PF speed mismatched the chart trend near station entry."
            entry_speed = _pf_chart_speed_kmph(chart_points[window_end]) if window_end < len(chart_points) else None
            if entry_speed is not None:
                pre_entry_window = [
                    _pf_chart_speed_kmph(point)
                    for point in chart_points[max(0, window_end - 12) : window_end + 1]
                ]
                pre_entry_window = [speed for speed in pre_entry_window if speed is not None]
                pre_entry_peak = max(pre_entry_window) if pre_entry_window else None
                zero_collapse_window = [
                    _pf_chart_speed_kmph(point)
                    for point in chart_points[window_end : min(len(chart_points), window_end + 36)]
                ]
                zero_collapse_window = [speed for speed in zero_collapse_window if speed is not None]
                if len(zero_collapse_window) >= 8:
                    zero_run = 0
                    longest_zero_run = 0
                    first_zero_index: int | None = None
                    pf_distance = _pf_speed_value(row.get("pf_distance"))
                    near_entry_window = zero_collapse_window[: min(10, len(zero_collapse_window))]
                    near_entry_peak = max(near_entry_window) if near_entry_window else None
                    sharp_post_entry_rise = any(
                        (near_entry_window[idx] - near_entry_window[idx - 1]) >= 6
                        for idx in range(1, len(near_entry_window))
                    ) if len(near_entry_window) >= 2 else False
                    for idx, speed in enumerate(zero_collapse_window):
                        if speed <= 5:
                            zero_run += 1
                            longest_zero_run = max(longest_zero_run, zero_run)
                            if first_zero_index is None:
                                first_zero_index = idx
                        else:
                            zero_run = 0
                    if (
                        pf_speed >= max(40.0, threshold)
                        and geofence_speed is not None
                        and near_entry_peak is not None
                        and near_entry_peak >= pf_speed - 1
                        and (
                            near_entry_peak >= entry_speed + 6
                            or sharp_post_entry_rise
                        )
                        and (pf_speed - geofence_speed) >= 8
                        and entry_speed <= pf_speed - 3.5
                        and first_zero_index is not None
                        and first_zero_index <= 8
                        and longest_zero_run >= 5
                    ):
                        return "Entry speed collapsed to zero too quickly after a local spike."
                    if (
                        pf_distance is not None
                        and 240 <= pf_distance <= 320
                        and pf_speed >= max(40.0, threshold)
                        and geofence_speed is not None
                        and near_entry_peak is not None
                        and near_entry_peak >= pf_speed - 1
                        and (
                            near_entry_peak >= entry_speed + 6
                            or sharp_post_entry_rise
                        )
                        and (pf_speed - geofence_speed) >= 8
                        and entry_speed <= pf_speed - 3.5
                        and first_zero_index is not None
                        and first_zero_index <= 20
                        and longest_zero_run >= 5
                    ):
                        return "Entry speed fell to zero shortly after a 250m-300m local spike."
                    if (
                        pf_distance is not None
                        and 240 <= pf_distance <= 320
                        and pf_speed >= max(40.0, threshold)
                        and geofence_speed is not None
                        and abs(pf_speed - geofence_speed) <= 6
                        and near_entry_peak is not None
                        and near_entry_peak >= pf_speed - 1
                        and entry_speed >= geofence_speed - 3
                        and entry_speed <= pf_speed + 2
                        and first_zero_index is not None
                        and first_zero_index <= 24
                        and longest_zero_run >= 10
                    ):
                        return "Entry speed stayed high only briefly before a sustained zero collapse near 250m-300m."
                    if (
                        pf_distance is not None
                        and 240 <= pf_distance <= 320
                        and stop_time_seconds is not None
                        and stop_time_seconds <= 45
                        and geofence_speed is not None
                        and pre_entry_peak is not None
                        and pre_entry_peak >= max(pf_speed, geofence_speed) + 8
                        and entry_speed >= min(pf_speed, geofence_speed) - 2
                        and first_zero_index is not None
                        and first_zero_index <= 18
                        and longest_zero_run >= 5
                    ):
                        return "Short-stop entry showed a sharp local peak before collapsing to zero near 250m-300m."
                    if (
                        pf_distance is not None
                        and 240 <= pf_distance <= 320
                        and stop_time_seconds is not None
                        and stop_time_seconds <= 240
                        and stop_time_seconds >= 46
                        and geofence_speed is not None
                        and pre_entry_peak is not None
                        and pre_entry_peak >= max(pf_speed, geofence_speed) + 2
                        and entry_speed >= min(pf_speed, geofence_speed) - 2
                        and first_zero_index is not None
                        and first_zero_index <= 30
                        and longest_zero_run >= 8
                    ):
                        return "Platform entry held a local peak before a sustained zero collapse near 250m-300m."
                    if (
                        pf_distance is not None
                        and 240 <= pf_distance <= 320
                        and stop_time_seconds is not None
                        and stop_time_seconds <= 420
                        and stop_time_seconds >= 241
                        and geofence_speed is not None
                        and pre_entry_peak is not None
                        and pre_entry_peak >= max(pf_speed, geofence_speed) + 15
                        and entry_speed >= min(pf_speed, geofence_speed) - 4
                        and first_zero_index is not None
                        and first_zero_index <= 20
                        and longest_zero_run >= 8
                    ):
                        return "Platform approach showed a strong local peak before a longer zero collapse near 250m-300m."

        pre_entry_speeds = [
            _pf_chart_speed_kmph(point)
            for point in chart_points[window_start : window_end + 1]
        ]
        pre_entry_speeds = [speed for speed in pre_entry_speeds if speed is not None]
        entry_speed = _pf_chart_speed_kmph(chart_points[window_end]) if window_end < len(chart_points) else None
        pf_distance = _pf_speed_value(row.get("pf_distance"))
        if entry_speed is not None and geofence_speed is not None:
            local_window_start = max(0, window_end - 12)
            local_window_speeds = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[local_window_start : window_end + 1]
            ]
            local_window_speeds = [speed for speed in local_window_speeds if speed is not None]
            if local_window_speeds:
                local_peak = max(local_window_speeds)
                if (
                    pf_speed >= max(45.0, threshold)
                    and stop_time_seconds is not None
                    and stop_time_seconds <= 90
                    and pf_distance is not None
                    and 240 <= pf_distance <= 320
                    and abs(entry_speed - geofence_speed) <= 8
                    and entry_speed <= 15
                    and local_peak >= pf_speed - 8
                    and (local_peak - entry_speed) >= 20
                    and (pf_speed - geofence_speed) >= 20
                ):
                    return "Stopped train had a 250m-300m pre-stop spike, likely network/GPS noise."
        if stop_time_seconds is not None and stop_time_seconds <= 90:
            pf_distance = _pf_speed_value(row.get("pf_distance"))
            station_window_start = max(0, window_end - 15)
            station_window_limit = min(
                len(chart_points) - 1,
                max(window_end, end_pos if end_pos is not None and end_pos >= 0 else window_end) + 6,
            )
            station_window_speeds = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[station_window_start : station_window_limit + 1]
            ]
            station_window_speeds = [speed for speed in station_window_speeds if speed is not None]
            if len(station_window_speeds) >= 6 and pf_distance is not None and 240 <= pf_distance <= 320:
                deltas = [
                    station_window_speeds[idx] - station_window_speeds[idx - 1]
                    for idx in range(1, len(station_window_speeds))
                ]
                sharp_rise = any(delta >= 8 for delta in deltas)
                sharp_drop = any(delta <= -10 for delta in deltas)
                direction_flips = sum(
                    1
                    for idx in range(1, len(deltas))
                    if abs(deltas[idx - 1]) >= 4
                    and abs(deltas[idx]) >= 4
                    and ((deltas[idx - 1] > 0 > deltas[idx]) or (deltas[idx - 1] < 0 < deltas[idx]))
                )
                local_peak = max(station_window_speeds)
                post_entry_window = station_window_speeds[-min(8, len(station_window_speeds)) :]
                post_entry_floor = min(post_entry_window) if post_entry_window else min(station_window_speeds)
                if (
                    pf_speed >= max(40.0, threshold)
                    and local_peak >= pf_speed - 2
                    and post_entry_floor <= 15
                    and (local_peak - post_entry_floor) >= 20
                    and sharp_drop
                    and (sharp_rise or direction_flips >= 1)
                ):
                    return "Sharp pre-stop chart swing near 250m-300m, likely network/GPS spike."
        if stop_time_seconds is not None and stop_time_seconds <= 90:
            pf_distance = _pf_speed_value(row.get("pf_distance"))
            if pf_distance is not None and 240 <= pf_distance <= 320:
                near_stop_end = min(
                    len(chart_points) - 1,
                    max(window_end, end_pos if end_pos is not None and end_pos >= 0 else window_end) + 2,
                )
                pre_stop_window = [
                    _pf_chart_speed_kmph(point)
                    for point in chart_points[max(0, window_end - 14) : window_end + 1]
                ]
                pre_stop_window = [speed for speed in pre_stop_window if speed is not None]
                stop_zone_window = [
                    _pf_chart_speed_kmph(point)
                    for point in chart_points[window_end : near_stop_end + 1]
                ]
                stop_zone_window = [speed for speed in stop_zone_window if speed is not None]
                entry_speed = _pf_chart_speed_kmph(chart_points[window_end]) if window_end < len(chart_points) else None
                if pre_stop_window and stop_zone_window and entry_speed is not None:
                    pre_stop_peak = max(pre_stop_window)
                    stop_zone_floor = min(stop_zone_window)
                    stop_zone_zero_run = 0
                    stop_zone_longest_zero_run = 0
                    stop_zone_first_zero_index: int | None = None
                    for idx, speed in enumerate(stop_zone_window):
                        if speed <= 5:
                            stop_zone_zero_run += 1
                            stop_zone_longest_zero_run = max(stop_zone_longest_zero_run, stop_zone_zero_run)
                            if stop_zone_first_zero_index is None:
                                stop_zone_first_zero_index = idx
                        else:
                            stop_zone_zero_run = 0
                    if (
                        pf_speed >= max(40.0, threshold)
                        and stop_zone_floor <= 5
                        and entry_speed <= min(20.0, pf_speed - 18)
                        and pre_stop_peak >= pf_speed - 2
                    ):
                        return "PF speed stayed high in the report, but charted entry collapsed before the stop."
                    if (
                        geofence_speed is not None
                        and pf_speed >= max(40.0, threshold)
                        and stop_zone_floor <= 5
                        and abs(entry_speed - geofence_speed) <= 6
                        and pre_stop_peak >= max(entry_speed, geofence_speed, pf_speed) + 12
                    ):
                        return "Chart showed a sharp local spike just before the stop window."
                    if (
                        geofence_speed is not None
                        and pf_speed >= max(40.0, threshold)
                        and stop_zone_first_zero_index is not None
                        and stop_zone_first_zero_index <= 20
                        and stop_zone_longest_zero_run >= 5
                        and entry_speed <= pf_speed - 4
                        and (pf_speed - geofence_speed) >= 8
                        and pre_stop_peak >= pf_speed - 1
                        and (
                            any(
                                (stop_zone_window[idx] - stop_zone_window[idx - 1]) >= 6
                                for idx in range(1, min(8, len(stop_zone_window)))
                            )
                            or max(stop_zone_window[: min(8, len(stop_zone_window))]) >= entry_speed + 6
                        )
                    ):
                        return "Entry speed collapsed to zero across the stop window after a short local spike."
                    initial_stop_zone = stop_zone_window[: min(6, len(stop_zone_window))]
                    if initial_stop_zone:
                        initial_zone_peak = max(initial_stop_zone)
                        if (
                            max(pf_speed, geofence_speed or 0) >= max(45.0, threshold)
                            and initial_zone_peak >= 60
                            and stop_zone_floor <= 5
                            and (initial_zone_peak - stop_zone_floor) >= 25
                            and (
                                (geofence_speed is not None and geofence_speed >= pf_speed + 15)
                                or initial_zone_peak >= pf_speed + 15
                            )
                        ):
                            return "Station entry started with an unusually high speed and collapsed too quickly."
                volatility_window_start = max(0, window_end - 10)
                volatility_window_end = min(
                    len(chart_points) - 1,
                    max(window_end + 6, end_pos if end_pos is not None and end_pos >= 0 else window_end),
                )
                volatility_window = [
                    _pf_chart_speed_kmph(point)
                    for point in chart_points[volatility_window_start : volatility_window_end + 1]
                ]
                volatility_window = [speed for speed in volatility_window if speed is not None]
                if len(volatility_window) >= 8:
                    deltas = [
                        volatility_window[idx] - volatility_window[idx - 1]
                        for idx in range(1, len(volatility_window))
                    ]
                    direction_flips = sum(
                        1
                        for idx in range(1, len(deltas))
                        if abs(deltas[idx - 1]) >= 4
                        and abs(deltas[idx]) >= 4
                        and ((deltas[idx - 1] > 0 > deltas[idx]) or (deltas[idx - 1] < 0 < deltas[idx]))
                    )
                    strong_rises = sum(1 for delta in deltas if delta >= 6)
                    strong_drops = sum(1 for delta in deltas if delta <= -6)
                    window_peak = max(volatility_window)
                    terminal_window = volatility_window[-min(10, len(volatility_window)) :]
                    window_floor = min(terminal_window)
                    near_zero_count = sum(1 for speed in terminal_window if speed <= 5)
                    if (
                        pf_speed >= max(40.0, threshold)
                        and window_peak >= pf_speed - 1
                        and window_floor <= 10
                        and (window_peak - window_floor) >= 20
                        and direction_flips >= 2
                        and strong_rises >= 2
                        and strong_drops >= 2
                        and near_zero_count >= 2
                    ):
                        return "Volatile pre-stop oscillation near station entry, likely network/GPS spike."
        if stop_time_seconds is not None and stop_time_seconds <= 180:
            pre_station_window = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[max(0, window_end - 8) : window_end + 1]
            ]
            pre_station_window = [speed for speed in pre_station_window if speed is not None]
            station_zone_end = min(
                len(chart_points) - 1,
                max(spike_window_end, end_pos if end_pos is not None and end_pos >= 0 else spike_window_end) + 4,
            )
            station_zone_window = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[window_end : station_zone_end + 1]
            ]
            station_zone_window = [speed for speed in station_zone_window if speed is not None]
            post_station_window = [
                _pf_chart_speed_kmph(point)
                for point in chart_points[
                    min(len(chart_points) - 1, max(window_end, end_pos if end_pos is not None and end_pos >= 0 else window_end)) :
                    min(len(chart_points), station_zone_end + 13)
                ]
            ]
            post_station_window = [speed for speed in post_station_window if speed is not None]
            if pre_station_window and station_zone_window and post_station_window:
                pre_station_peak = max(pre_station_window)
                station_zone_floor = min(station_zone_window)
                post_station_peak = max(post_station_window)
                station_zero_run = 0
                station_longest_zero_run = 0
                for speed in station_zone_window:
                    if speed <= 5:
                        station_zero_run += 1
                        station_longest_zero_run = max(station_longest_zero_run, station_zero_run)
                    else:
                        station_zero_run = 0
                if (
                    pf_speed >= max(40.0, threshold)
                    and pre_station_peak >= max(45.0, threshold)
                    and post_station_peak >= max(40.0, threshold)
                    and station_zone_floor <= 5
                    and station_longest_zero_run >= 3
                    and (pre_station_peak - station_zone_floor) >= 22
                    and (post_station_peak - station_zone_floor) >= 22
                ):
                    return "Station window dropped into a zero pocket and rebounded quickly, likely GPS/network spike."
        if len(pre_entry_speeds) >= 8:
            peak_speed = max(pre_entry_speeds)
            peak_index = pre_entry_speeds.index(peak_speed)
            post_peak = pre_entry_speeds[peak_index:]
            upward_bursts = sum(
                1
                for idx in range(1, len(post_peak))
                if (post_peak[idx] - post_peak[idx - 1]) > 4
            )
            if (
                peak_speed >= max(45.0, threshold)
                and len(post_peak) >= 5
                and post_peak[-1] <= peak_speed - 20
                and upward_bursts <= max(1, len(post_peak) // 6)
            ):
                return None

    speed_400m = _pf_speed_value(row.get("speed_at_400m"))
    speed_265m = _pf_speed_value(row.get("speed_at_265m"))
    speed_100m = _pf_speed_value(row.get("speed_at_100m"))
    if (
        speed_400m is not None
        and speed_265m is not None
        and speed_100m is not None
        and speed_400m >= speed_265m >= speed_100m
        and (speed_400m - speed_100m) >= 15
        and pf_speed < 90
    ):
        return None

    # Extremely high PF/geofence speeds at a station with a very short stop
    # are almost always GPS/network spikes in this workflow.
    if max_entry_speed >= 100 and stop_time_seconds is not None and stop_time_seconds <= 60:
        return "Very high entry speed with a very short stop, likely GPS/network spike."

    previous_speed = _pf_reference_speed(previous_row)
    next_speed = _pf_reference_speed(next_row)

    reference_speeds = [speed for speed in (geofence_speed, previous_speed, next_speed) if speed is not None]
    if len(reference_speeds) < 2:
        return None

    large_gap_count = sum(1 for speed in reference_speeds if (pf_speed - speed) >= 15)
    severe_gap_count = sum(1 for speed in reference_speeds if (pf_speed - speed) >= 25)
    isolated_peak = (
        previous_speed is not None
        and next_speed is not None
        and (pf_speed - max(previous_speed, next_speed)) >= 25
    )
    geofence_mismatch = geofence_speed is not None and abs(pf_speed - geofence_speed) >= 25

    # For moderate 40-60 type values, avoid auto-omitting based on neighboring
    # station summaries alone. Those cases need the chart trend for confidence.
    if pf_speed < 80:
        return None

    if severe_gap_count >= 2 and (isolated_peak or geofence_mismatch or large_gap_count >= 3):
        return "PF speed was isolated from neighboring station/reference speeds."
    return None


def _pf_is_suspected_spike(
    row: dict[str, object],
    previous_row: dict[str, object] | None,
    next_row: dict[str, object] | None,
    threshold: int,
    chart_points: list[dict[str, object]] | None = None,
) -> bool:
    return _pf_suspected_spike_reason(row, previous_row, next_row, threshold, chart_points) is not None


def _pf_row_signature(row: dict[str, object]) -> tuple[str, str, str, str, str, str]:
    return (
        str(row.get("train_no") or "").strip(),
        str(row.get("station") or "").strip(),
        str(row.get("srl_no") or "").strip(),
        str(row.get("act_arr") or "").strip(),
        str(row.get("act_dep") or "").strip(),
        str(row.get("pf_enter_speed") or "").strip(),
    )


def _pf_run_level_spike_reason(
    rows: list[dict[str, object]],
    chart_points: list[dict[str, object]] | None,
) -> str | None:
    if not chart_points or len(chart_points) < 80:
        return None

    speed_values = [
        0.0 if speed is None else float(speed)
        for speed in (_pf_chart_speed_kmph(point) for point in chart_points)
    ]
    fast_rebound_count = 0
    fast_collapse_count = 0
    hard_reset_count = 0
    noisy_window_count = 0
    for idx in range(len(speed_values) - 8):
        current_speed = speed_values[idx]
        upcoming_window = speed_values[idx + 1 : idx + 9]
        if current_speed <= 5 and max(upcoming_window, default=0.0) >= 45:
            fast_rebound_count += 1
        if current_speed >= 45 and min(upcoming_window, default=999.0) <= 5:
            fast_collapse_count += 1
        if (
            (current_speed <= 5 and max(upcoming_window[:5], default=0.0) >= 35)
            or (current_speed >= 35 and min(upcoming_window[:5], default=999.0) <= 5)
        ):
            hard_reset_count += 1
        local_span = max(upcoming_window, default=current_speed) - min(upcoming_window, default=current_speed)
        if current_speed <= 12 and max(upcoming_window[:6], default=0.0) >= 40 and local_span >= 28:
            noisy_window_count += 1

    candidate_rows = 0
    short_mismatch_rows = 0
    short_stop_rows = 0
    for row in rows:
        stop_time_seconds = _parse_hms_seconds(row.get("stop_time"))
        pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
        geofence_speed = _pf_speed_value(row.get("geofence_enter_speed"))
        max_speed = max((speed for speed in (pf_speed, geofence_speed) if speed is not None), default=None)
        if stop_time_seconds is None or stop_time_seconds == 0 or max_speed is None or max_speed <= 40:
            continue
        candidate_rows += 1
        if stop_time_seconds <= 60:
            short_stop_rows += 1
        pf_distance = _pf_speed_value(row.get("pf_distance"))
        if (
            pf_distance is not None
            and 240 <= pf_distance <= 320
            and stop_time_seconds <= 45
            and pf_speed is not None
            and geofence_speed is not None
            and abs(pf_speed - geofence_speed) >= 20
        ):
            short_mismatch_rows += 1

    if (
        fast_rebound_count >= 8
        and candidate_rows >= 4
        and (fast_collapse_count >= 3 or short_mismatch_rows >= 2)
    ):
        return "Train chart showed repeated zero-to-high rebounds across multiple stations."
    if (
        candidate_rows >= 3
        and short_stop_rows >= 2
        and fast_rebound_count >= 5
        and fast_collapse_count >= 5
        and (hard_reset_count >= 10 or noisy_window_count >= 4)
    ):
        return "Train chart showed repeated hard speed resets across multiple PF stops."
    if (
        candidate_rows >= 2
        and fast_rebound_count >= 4
        and fast_collapse_count >= 4
        and short_mismatch_rows >= 1
        and hard_reset_count >= 8
    ):
        return "Train chart showed repeated rebound/collapse noise near PF stops."
    if (
        fast_rebound_count >= 6
        and candidate_rows <= 2
        and short_mismatch_rows >= 1
    ):
        return "Train chart showed repeated rebound noise despite only a few PF events."
    return None


def _pf_should_omit_train_after_row_filter(
    rows: list[dict[str, object]],
    train_kept_rows: list[dict[str, object]],
    train_spike_rows: list[dict[str, object]],
    chart_points: list[dict[str, object]] | None,
) -> str | None:
    if not chart_points or len(chart_points) < 80 or not train_kept_rows or not train_spike_rows:
        return None

    speed_values = [
        0.0 if speed is None else float(speed)
        for speed in (_pf_chart_speed_kmph(point) for point in chart_points)
    ]
    fast_rebound_count = 0
    fast_collapse_count = 0
    hard_reset_count = 0
    noisy_window_count = 0
    for idx in range(len(speed_values) - 8):
        current_speed = speed_values[idx]
        upcoming_window = speed_values[idx + 1 : idx + 9]
        if current_speed <= 5 and max(upcoming_window, default=0.0) >= 45:
            fast_rebound_count += 1
        if current_speed >= 45 and min(upcoming_window, default=999.0) <= 5:
            fast_collapse_count += 1
        if (
            (current_speed <= 5 and max(upcoming_window[:5], default=0.0) >= 35)
            or (current_speed >= 35 and min(upcoming_window[:5], default=999.0) <= 5)
        ):
            hard_reset_count += 1
        local_span = max(upcoming_window, default=current_speed) - min(upcoming_window, default=current_speed)
        if current_speed <= 12 and max(upcoming_window[:6], default=0.0) >= 40 and local_span >= 28:
            noisy_window_count += 1

    if (
        len(train_spike_rows) >= 3
        and len(train_kept_rows) <= len(train_spike_rows)
        and fast_rebound_count >= 4
        and fast_collapse_count >= 4
        and (hard_reset_count >= 7 or noisy_window_count >= 3)
    ):
        return "Train still showed repeated chart resets after row filtering, so all remaining PF rows were omitted."
    if (
        len(train_spike_rows) >= 2
        and len(train_kept_rows) <= 2
        and hard_reset_count >= 9
        and noisy_window_count >= 3
    ):
        return "Remaining PF rows belonged to a chart with repeated spike/reset behavior, so the full train was omitted."
    return None


def _pf_manual_train_spike_reason(report_day: date, train_no: str) -> str | None:
    override_trains = SSTS_PF_SPIKE_FILTER_TRAIN_OVERRIDES.get(report_day.isoformat(), set())
    if train_no.strip() in override_trains:
        return "Train manually marked for spike omission after chart review."
    return None


def _normalize_ssts_crew_name(value: object | None) -> str:
    text = str(value or "").strip().upper()
    return re.sub(r"[^A-Z0-9]+", "", text)


def fetch_ssts_crew_lookup(token: str) -> dict[str, str]:
    global _SSTS_CREW_CACHE
    now_utc = _utc_now()
    if _SSTS_CREW_CACHE is not None:
        cached_at, cached_lookup = _SSTS_CREW_CACHE
        if (now_utc - cached_at) < timedelta(minutes=SSTS_PF_REPORT_CACHE_TTL_MINUTES):
            return dict(cached_lookup)

    response = _ssts_get_json(SSTS_API_CREW_URL, headers={"Authorization": token})
    if not isinstance(response, dict):
        raise RuntimeError("Unexpected SSTS crew response format.")
    raw_rows = response.get("data")
    if not isinstance(raw_rows, list):
        raise RuntimeError("Unexpected SSTS crew data payload.")

    lookup: dict[str, str] = {}
    for item in raw_rows:
        if not isinstance(item, dict):
            continue
        crew_name_key = _normalize_ssts_crew_name(item.get("crew_name"))
        crew_id = str(item.get("crew_id") or "").strip()
        if crew_name_key and crew_id and crew_name_key not in lookup:
            lookup[crew_name_key] = crew_id
    for crew_name, crew_id in SSTS_PF_CREW_ID_OVERRIDES.items():
        crew_name_key = _normalize_ssts_crew_name(crew_name)
        if crew_name_key and crew_id:
            lookup[crew_name_key] = crew_id

    _SSTS_CREW_CACHE = (now_utc, lookup)
    return dict(lookup)


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


def _build_ssts_pf_speed_analysis_result(
    report_day: date,
    speed_threshold: int,
    progress_callback: Callable[[int, str], None] | None = None,
) -> dict[str, object]:
    def report_progress(percent: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(percent, message)

    raw_context = build_ssts_pf_entering_context(report_day)
    token = fetch_ssts_token()
    report_progress(18, "Loaded PF source rows.")
    detailed_analysis_threshold = 40
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

    daily_report_rows: list[dict[str, object]] = []
    for row in raw_context.get("pf_report_rows", []):
        if not isinstance(row, dict):
            continue
        stop_time = str(row.get("stop_time") or "").strip()
        if stop_time == "00:00:00":
            continue
        pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
        if _pf_speed_matches_threshold(pf_speed, speed_threshold):
            daily_report_rows.append(dict(row))

    daily_report_rows.sort(
        key=lambda row: (
            str(row.get("train_no") or ""),
            999999 if row.get("srl_no") in ("", None) else int(row.get("srl_no") or 0),
            str(row.get("station") or ""),
        )
    )

    detailed_candidate_trains = {
        str(row.get("train_no") or "").strip()
        for row in raw_context.get("pf_report_rows", [])
        if isinstance(row, dict)
        and _pf_speed_matches_threshold(_pf_speed_value(row.get("pf_enter_speed")), detailed_analysis_threshold)
    }
    chart_points_by_train: dict[str, list[dict[str, object]]] = {train_no: [] for train_no in detailed_candidate_trains if train_no}
    candidate_rows_by_train = {
        train_no: rows
        for train_no, rows in all_rows_by_train.items()
        if train_no in detailed_candidate_trains and rows
    }
    if candidate_rows_by_train:
        total_candidates = len(candidate_rows_by_train)
        processed_candidates = 0
        report_progress(28, f"Fetching speed charts for {total_candidates} trains...")
        with ThreadPoolExecutor(max_workers=6) as executor:
            future_map = {
                executor.submit(_fetch_ssts_positions, rows[0], token): train_no
                for train_no, rows in candidate_rows_by_train.items()
            }
            for future in as_completed(future_map):
                train_no = future_map[future]
                try:
                    chart_points_by_train[train_no] = future.result()
                except (urlerror.URLError, RuntimeError, ValueError, json.JSONDecodeError):
                    chart_points_by_train[train_no] = []
                processed_candidates += 1
                chart_progress = 28 + int((processed_candidates / total_candidates) * 42)
                report_progress(
                    chart_progress,
                    f"Fetched speed charts for {processed_candidates}/{total_candidates} trains...",
                )

    detailed_daily_report_rows: list[dict[str, object]] = []
    suspected_spike_rows: list[dict[str, object]] = []
    suspected_spike_signatures: set[tuple[str, str, str, str, str, str]] = set()
    trains_to_review = list(all_rows_by_train.items())
    total_trains_to_review = len(trains_to_review)
    for train_index, (train_no, rows) in enumerate(trains_to_review, start=1):
        chart_points = chart_points_by_train.get(train_no, [])
        train_kept_rows: list[dict[str, object]] = []
        train_spike_rows: list[dict[str, object]] = []
        for index, row in enumerate(rows):
            stop_time = str(row.get("stop_time") or "").strip()
            if stop_time == "00:00:00":
                continue
            pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
            if not _pf_speed_matches_threshold(pf_speed, detailed_analysis_threshold):
                continue
            previous_row = rows[index - 1] if index > 0 else None
            next_row = rows[index + 1] if index + 1 < len(rows) else None
            spike_reason = _pf_suspected_spike_reason(
                row,
                previous_row,
                next_row,
                detailed_analysis_threshold,
                chart_points,
            )
            if spike_reason is not None:
                spike_row = dict(row)
                spike_row["spike_reason"] = spike_reason
                current_remarks = str(spike_row.get("remarks") or "").strip()
                spike_row["remarks"] = (
                    f"{current_remarks} | Omitted: {spike_reason}" if current_remarks else f"Omitted: {spike_reason}"
                )
                train_spike_rows.append(spike_row)
                continue
            if _pf_speed_matches_threshold(pf_speed, speed_threshold):
                train_kept_rows.append(dict(row))
        run_level_reason = _pf_manual_train_spike_reason(report_day, train_no)
        if run_level_reason is None:
            run_level_reason = _pf_run_level_spike_reason(rows, chart_points)
        if run_level_reason is None:
            run_level_reason = _pf_should_omit_train_after_row_filter(
                rows,
                train_kept_rows,
                train_spike_rows,
                chart_points,
            )
        if run_level_reason is None and len(train_spike_rows) >= 5 and train_kept_rows:
            run_level_reason = "Train showed repeated spike patterns across multiple PF stops."
        if run_level_reason and train_kept_rows:
            for kept_row in train_kept_rows:
                spike_row = dict(kept_row)
                spike_row["spike_reason"] = run_level_reason
                current_remarks = str(spike_row.get("remarks") or "").strip()
                spike_row["remarks"] = (
                    f"{current_remarks} | Omitted: {run_level_reason}" if current_remarks else f"Omitted: {run_level_reason}"
                )
                train_spike_rows.append(spike_row)
            train_kept_rows = []
        for spike_row in train_spike_rows:
            suspected_spike_rows.append(spike_row)
            suspected_spike_signatures.add(_pf_row_signature(spike_row))
        detailed_daily_report_rows.extend(train_kept_rows)
        if total_trains_to_review:
            review_progress = 72 + int((train_index / total_trains_to_review) * 22)
            report_progress(
                review_progress,
                f"Reviewing spike cases train-wise... {train_index}/{total_trains_to_review}",
            )

    detailed_daily_report_rows.sort(
        key=lambda row: (
            str(row.get("train_no") or ""),
            999999 if row.get("srl_no") in ("", None) else int(row.get("srl_no") or 0),
            str(row.get("station") or ""),
        )
    )
    detailed_detail_rows_by_train: dict[str, list[dict[str, object]]] = {}
    for train_no, rows in all_rows_by_train.items():
        cleaned_rows = [
            dict(row) for row in rows
            if _pf_row_signature(row) not in suspected_spike_signatures
        ]
        detailed_detail_rows_by_train[train_no] = cleaned_rows
    report_progress(96, "Finalizing detailed PF report...")

    return {
        "pf_report_day": report_day.isoformat(),
        "pf_report_day_label": report_day.strftime("%d-%m-%Y"),
        "pf_speed_threshold": speed_threshold,
        "pf_speed_threshold_label": _pf_speed_threshold_label(speed_threshold),
        "pf_daily_report_rows": daily_report_rows,
        "pf_detailed_daily_report_rows": detailed_daily_report_rows,
        "pf_detailed_daily_spike_count": len(suspected_spike_rows),
        "pf_detailed_daily_spike_rows": suspected_spike_rows,
        "pf_analysis_summary_rows": summary_rows,
        "pf_analysis_detail_rows_by_train": detail_rows_by_train,
        "pf_detailed_detail_rows_by_train": detailed_detail_rows_by_train,
        "pf_analysis_total_trains": len(summary_rows),
        "pf_analysis_total_rows": len(filtered_rows),
        "pf_analysis_source_total_trains": int(raw_context.get("pf_report_total_trains") or 0),
        "pf_analysis_source_total_rows": int(raw_context.get("pf_report_total_rows") or 0),
        "pf_analysis_missing_count": int(raw_context.get("pf_report_missing_count") or 0),
    }


def _run_ssts_pf_analysis_task(task_id: str, report_day: date, speed_threshold: int) -> None:
    try:
        def push_progress(percent: int, message: str) -> None:
            _set_ssts_pf_analysis_task(
                task_id,
                status="running",
                progress=percent,
                message=message,
            )

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
        result = _build_ssts_pf_speed_analysis_result(report_day, speed_threshold, push_progress)
        _set_ssts_pf_analysis_task(
            task_id,
            progress=99,
            message="Wrapping up PF analysis...",
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


def _prune_old_ssts_snapshots(
    session: Session,
    retention_days: int = SSTS_SNAPSHOT_RETENTION_DAYS,
) -> dict[str, int]:
    cutoff_utc = _utc_now() - timedelta(days=retention_days)
    old_runs = list(
        session.exec(
            select(SstsSnapshotRun).where(SstsSnapshotRun.observed_at <= cutoff_utc)
        ).all()
    )
    if not old_runs:
        return {"deleted_runs": 0, "deleted_snapshots": 0}

    run_ids = [run.id for run in old_runs if run.id is not None]
    deleted_snapshots = 0
    if run_ids:
        old_snapshots = list(
            session.exec(
                select(SstsDeviceSnapshot).where(SstsDeviceSnapshot.run_id.in_(run_ids))
            ).all()
        )
        deleted_snapshots = len(old_snapshots)
        for snapshot in old_snapshots:
            session.delete(snapshot)

    for run in old_runs:
        session.delete(run)

    session.commit()
    return {"deleted_runs": len(old_runs), "deleted_snapshots": deleted_snapshots}


def refresh_ssts_snapshot(session: Session, force: bool = False) -> dict[str, object]:
    cleanup_summary = _prune_old_ssts_snapshots(session)
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
                "cleanup_summary": cleanup_summary,
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
            "cleanup_summary": cleanup_summary,
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
            "cleanup_summary": cleanup_summary,
        }


def _run_background_ssts_sync_once(force: bool = False) -> None:
    with Session(engine) as session:
        refresh_ssts_snapshot(session, force=force)


def _background_ssts_sync_worker() -> None:
    while not _SSTS_BACKGROUND_SYNC_STOP.is_set():
        try:
            _run_background_ssts_sync_once(force=False)
        except Exception:
            pass
        wait_seconds = max(60, SSTS_BACKGROUND_SYNC_INTERVAL_MINUTES * 60)
        if _SSTS_BACKGROUND_SYNC_STOP.wait(wait_seconds):
            break


def _ensure_background_ssts_sync() -> None:
    global _SSTS_BACKGROUND_SYNC_THREAD
    if _SSTS_BACKGROUND_SYNC_THREAD and _SSTS_BACKGROUND_SYNC_THREAD.is_alive():
        return
    _SSTS_BACKGROUND_SYNC_STOP.clear()
    _SSTS_BACKGROUND_SYNC_THREAD = threading.Thread(
        target=_background_ssts_sync_worker,
        name="ssts-background-sync",
        daemon=True,
    )
    _SSTS_BACKGROUND_SYNC_THREAD.start()


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
    selected_analysis_rake: int | None = None,
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
            "selected_analysis_rake": None,
            "selected_analysis_rake_label": "",
            "selected_analysis_history_rows": [],
            "selected_analysis_history_summary": None,
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
    selected_analysis_history_rows: list[dict[str, object]] = []
    selected_analysis_rake_value: str | None = None
    selected_analysis_rake_label = ""
    selected_analysis_history_summary: dict[str, object] | None = None
    selected_analysis_summary = {
        "day_label": analysis_day_value.strftime("%d-%m-%Y") if analysis_day_value else "",
        "continuous_offline_count": 0,
        "continuous_online_count": 0,
        "mixed_online_offline_count": 0,
        "online_offline_total_count": 0,
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

            if seed_state is not None and points[0]["time"] > day_start:
                points.insert(0, {"time": day_start, "state": seed_state})

            segments: list[dict[str, object]] = []

            def build_segment(
                state: str,
                start_time: datetime,
                end_time: datetime,
                *,
                count_for_periods: bool = False,
            ) -> dict[str, object] | None:
                if end_time <= start_time:
                    return None
                duration_minutes = max(0, int((end_time - start_time).total_seconds() // 60))
                display_end = end_time - timedelta(minutes=1)
                return {
                    "state": state,
                    "start_time": start_time,
                    "end_time": end_time,
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
                    segment = build_segment(
                        current_state,
                        segment_start,
                        point_time,
                        count_for_periods=duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                    )
                    if segment is not None:
                        segments.append(segment)
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

            offline_transition_time = _ssts_offline_transition_time(lastupdate_time)
            if offline_transition_time is not None:
                bounded_offline_start = min(reference_end, max(day_start, offline_transition_time))
                if current_state == "offline":
                    current_segment_start = min(current_segment_start, bounded_offline_start)
                elif current_state == "online" and current_segment_start < bounded_offline_start < reference_end:
                    current_segment_end = bounded_offline_start
                    followup_offline_start = bounded_offline_start

            duration_minutes = max(0, int((current_segment_end - current_segment_start).total_seconds() // 60))
            if duration_minutes > 0:
                display_end_time = current_segment_end
                if current_state == "online" and offline_transition_time is not None:
                    display_end_time = min(current_segment_end, max(current_segment_start, offline_transition_time))
                segment = build_segment(
                    current_state,
                    current_segment_start,
                    display_end_time,
                    count_for_periods=duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                )
                if segment is not None:
                    segments.append(segment)
                elif current_state == "online":
                    followup_offline_start = current_segment_start

            if followup_offline_start is not None and followup_offline_start < reference_end:
                offline_duration_minutes = max(0, int((reference_end - followup_offline_start).total_seconds() // 60))
                if offline_duration_minutes > 0:
                    segment = build_segment(
                        "offline",
                        followup_offline_start,
                        reference_end,
                        count_for_periods=offline_duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                    )
                    if segment is not None:
                        segments.append(segment)

            normalized_segments: list[dict[str, object]] = []
            for segment in sorted(segments, key=lambda item: item["start_time"]):
                if not normalized_segments:
                    normalized_segments.append(segment)
                    continue
                previous_segment = normalized_segments[-1]
                if (
                    previous_segment["state"] == segment["state"]
                    and segment["start_time"] <= previous_segment["end_time"]
                ):
                    merged_end = max(previous_segment["end_time"], segment["end_time"])
                    merged_duration_minutes = max(
                        0,
                        int((merged_end - previous_segment["start_time"]).total_seconds() // 60),
                    )
                    previous_segment["end_time"] = merged_end
                    previous_segment["end_label"] = _format_ist_time(merged_end - timedelta(minutes=1))
                    previous_segment["duration_label"] = _format_duration(merged_duration_minutes) or ""
                    previous_segment["width_percent"] = round((merged_duration_minutes / (24 * 60)) * 100, 2)
                    previous_segment["summary_label"] = (
                        f"{previous_segment['start_label']} to {previous_segment['end_label']} "
                        f"{'offline' if previous_segment['state'] == 'offline' else 'online'}"
                    )
                    previous_segment["count_for_periods"] = (
                        merged_duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES
                    )
                    continue
                normalized_segments.append(segment)
            segments = normalized_segments

            previous_segment_end = day_start
            for segment in segments:
                gap_minutes = max(
                    0,
                    int((segment["start_time"] - previous_segment_end).total_seconds() // 60),
                )
                segment["gap_before_percent"] = round((gap_minutes / (24 * 60)) * 100, 2)
                previous_segment_end = max(previous_segment_end, segment["end_time"])

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
        selected_analysis_summary["continuous_online_count"] = sum(
            1
            for row in selected_analysis_rows
            if row.get("segments")
            and {str(segment.get("state")) for segment in row.get("segments", [])} == {"online"}
        )
        selected_analysis_summary["mixed_online_offline_count"] = sum(
            1
            for row in selected_analysis_rows
            if {"online", "offline"}.issubset({str(segment.get("state")) for segment in row.get("segments", [])})
        )
        selected_analysis_summary["online_offline_total_count"] = (
            int(selected_analysis_summary["continuous_online_count"])
            + int(selected_analysis_summary["mixed_online_offline_count"])
        )

        selected_analysis_rows.sort(
            key=lambda item: (
                -int(item.get("offline_periods") or 0),
                -int(item.get("online_periods") or 0),
                str(item.get("name") or "").lower(),
            )
        )

        selected_rake_row = next(
            (
                row
                for row in selected_analysis_rows
                if selected_analysis_rake is not None and int(row.get("device_id") or 0) == selected_analysis_rake
            ),
            None,
        )
        if selected_rake_row is not None:
            selected_analysis_rake_value = str(selected_analysis_rake)
            selected_analysis_rake_label = str(selected_rake_row.get("name") or "")
            for run in selected_day_runs:
                run_time = _ensure_utc(run.observed_at)
                if run_time is None:
                    continue
                matching_snapshot = next(
                    (
                        snapshot
                        for snapshot in day_rows_by_run.get(run.id or 0, [])
                        if snapshot.device_id == selected_analysis_rake
                    ),
                    None,
                )
                if matching_snapshot is None:
                    continue
                snapshot_state = "offline" if _ssts_is_offline(matching_snapshot) else "online"
                snapshot_offline_minutes = _snapshot_offline_minutes(matching_snapshot, run_time)
                selected_analysis_history_rows.append(
                    {
                        "observed_at_label": _format_ist(run_time, include_seconds=True),
                        "status": "Offline" if snapshot_state == "offline" else "Online",
                        "status_class": snapshot_state,
                        "lastupdate_label": _format_ist(matching_snapshot.lastupdate, include_seconds=True),
                        "offline_duration": _format_duration(snapshot_offline_minutes) or "",
                        "remark": matching_snapshot.remark or "",
                    }
                )

            selected_analysis_history_summary = {
                "selected_count": len(selected_analysis_history_rows),
                "online_count": sum(1 for row in selected_analysis_history_rows if row["status_class"] == "online"),
                "offline_count": sum(1 for row in selected_analysis_history_rows if row["status_class"] == "offline"),
                "timeline_summary": str(selected_rake_row.get("timeline_summary") or ""),
            }

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
        "selected_analysis_rake": selected_analysis_rake_value,
        "selected_analysis_rake_label": selected_analysis_rake_label,
        "selected_analysis_history_rows": selected_analysis_history_rows,
        "selected_analysis_history_summary": selected_analysis_history_summary,
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
    _ensure_background_ssts_sync()


@app.on_event("shutdown")
def on_shutdown() -> None:
    _SSTS_BACKGROUND_SYNC_STOP.set()


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
        response.set_cookie(_AUTH_COOKIE, "ok", httponly=True, max_age=86400, path="/", samesite="lax")
        return response
    return templates.TemplateResponse(
        "login.html",
        {"request": request, "error": "Invalid credentials"},
        status_code=401,
    )


@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie(_AUTH_COOKIE, path="/", samesite="lax")
    return response


def _safe_return_to(value: str | None, fallback: str = "/employees") -> str:
    target = (value or "").strip()
    if not target:
        return fallback
    if not target.startswith("/"):
        return fallback
    if target.startswith("//"):
        return fallback
    return target


def _retired_employee_candidates(session: Session, *, today_value: date | None = None) -> list[Employee]:
    today_value = today_value or date.today()
    employees = session.exec(select(Employee).where(Employee.retirement_date.is_not(None))).all()
    return sorted(
        [employee for employee in employees if employee.retirement_date and employee.retirement_date < today_value],
        key=lambda employee: (employee.retirement_date or date.min, role_sort_key(employee.role), employee.name),
    )


def _is_hidden_employee_role(role: object | None) -> bool:
    normalized = normalize_role(_clean_import_text(role) or "")
    return normalized in HIDDEN_EMPLOYEE_ROLES


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
    _sync_retired_employees(session, today)
    _sync_employee_duplicates(session)
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
    role_headers, working_summary = build_working_location_summary(employees_now)

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
            "role_headers": role_headers,
            "working_summary": working_summary,
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
    category: Optional[str] = None,
    working_at: Optional[str] = None,
    cli: Optional[str] = None,
    cli_status: Optional[str] = None,
    gradation: Optional[str] = None,
    sort: str = "role",
    page: int = 1,
    per_page: int = 100,
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_gradation: Optional[str] = None,
    sync_notice: Optional[str] = None,
    sync_warning: Optional[str] = None,
    sync_error: Optional[str] = None,
    retired_cleanup_notice: Optional[str] = None,
    retired_cleanup_error: Optional[str] = None,
    retired_preview: Optional[str] = None,
    session: Session = Depends(get_session),
):
    today = date.today()
    _sync_retired_employees(session, today)
    _sync_employee_duplicates(session)
    roster_filter_active = any([roster_name, roster_cli, roster_gradation])
    employees_open = not roster_filter_active
    roster_open = roster_filter_active

    employees_all = fetch_active_employees(session, today)
    raw_working = {e.working_at for e in employees_all if e.working_at}
    working_opts_filtered = {wa for wa in raw_working if wa.upper().startswith("CC(")}
    working_opts = sorted(working_opts_filtered if working_opts_filtered else raw_working)
    cli_opts_map: dict[str, str] = {}
    for cli_name, cli_id in session.exec(select(Employee.cli, Employee.cli_id).distinct()):
        val = format_cli_label(cli_name, cli_id)
        if not val:
            continue
        key = val.strip().lower()
        if key not in cli_opts_map:
            cli_opts_map[key] = val.strip()
    cli_opts = [v for _, v in sorted(cli_opts_map.items(), key=lambda item: item[0])]
    category_opts = sorted({value for value in session.exec(select(Employee.category).distinct()) if value})
    gradation_opts = sorted({value for value in session.exec(select(Employee.gradation).distinct()) if value})

    page = 1
    per_page = 0

    query_employees = select(Employee)
    if role:
        query_employees = query_employees.where(Employee.role == role)
    if category:
        category_lower = category.strip().lower()
        query_employees = query_employees.where(func.lower(Employee.category) == category_lower)
    if working_at:
        query_employees = query_employees.where(Employee.working_at.ilike(f"%{working_at}%"))
    if cli_status == "assigned":
        query_employees = query_employees.where(func.trim(func.coalesce(Employee.cli, "")) != "")
    elif cli_status == "unassigned":
        query_employees = query_employees.where(func.trim(func.coalesce(Employee.cli, "")) == "")
    if gradation:
        query_employees = query_employees.where(Employee.gradation.ilike(f"%{gradation}%"))

    total_count = 0
    employees: list[Employee] = []
    if q or cli or cli_status:
        employees_all = session.exec(query_employees).all()
        employees = [e for e in employees_all if not _is_hidden_employee_role(e.role)]
        if q:
            q_lower = q.lower()
            employees = [
                e
                for e in employees
                if q_lower in e.name.lower()
            ]
        if cli:
            cli_lower = cli.strip().lower()
            employees = [e for e in employees if cli_lower in _employee_cli_label(e).lower()]

        def sort_key(e: Employee):
            if sort == "name":
                return (e.name.lower(),)
            if sort == "retirement":
                return (e.retirement_date or date.max, e.name)
            if sort == "category":
                return ((e.category or "").lower(), e.name.lower())
            if sort == "gradation":
                return ((e.gradation or "").lower(), e.name.lower())
            if sort == "hire":
                return (e.hire_date, e.name)
            if sort == "cli":
                return (_employee_cli_key(e), e.name)
            if sort == "working_at":
                return ((e.working_at or "").lower(), e.name)
            return (role_sort_key(e.role), e.name)

        employees = sorted(employees, key=sort_key)
        total_count = len(employees)
    else:
        role_case = case(
            {role_name: idx for idx, role_name in enumerate(ROLE_ORDER)},
            value=Employee.role,
            else_=999,
        )
        if sort == "name":
            query_employees = query_employees.order_by(Employee.name)
        elif sort == "retirement":
            query_employees = query_employees.order_by(Employee.retirement_date, Employee.name)
        elif sort == "category":
            query_employees = query_employees.order_by(Employee.category, Employee.name)
        elif sort == "gradation":
            query_employees = query_employees.order_by(Employee.gradation, Employee.name)
        elif sort == "hire":
            query_employees = query_employees.order_by(Employee.hire_date, Employee.name)
        elif sort == "working_at":
            query_employees = query_employees.order_by(Employee.working_at, Employee.name)
        elif sort == "cli":
            query_employees = query_employees.order_by(Employee.cli, Employee.name)
        else:
            query_employees = query_employees.order_by(role_case, Employee.name)

        employees = [e for e in session.exec(query_employees).all() if not _is_hidden_employee_role(e.role)]
        total_count = len(employees)

    cli_roster: list[Employee] = []
    if roster_filter_active:
        if not employees_all:
            employees_all = session.exec(select(Employee)).all()
        cli_roster = [e for e in employees_all if e.cli and not _is_hidden_employee_role(e.role)]
    if roster_name:
        name_lower = roster_name.lower()
        cli_roster = [e for e in cli_roster if name_lower in e.name.lower()]
    if roster_cli:
        roster_cli_lower = roster_cli.strip().lower()
        cli_roster = [e for e in cli_roster if roster_cli_lower in _employee_cli_label(e).lower()]
    if roster_gradation:
        grad_lower = roster_gradation.lower()
        cli_roster = [e for e in cli_roster if e.gradation and grad_lower in e.gradation.lower()]
    cli_roster = sorted(cli_roster, key=lambda e: (_employee_cli_key(e), e.name))

    total_pages = 1
    page_start = 0
    prev_url = ""
    next_url = ""
    employee_return_to = f"{request.url.path}{('?' + request.url.query) if request.url.query else ''}#employees-card"
    employee_return_to_query = urlparse.quote(employee_return_to, safe="/")
    retired_preview_rows = _retired_employee_candidates(session) if retired_preview == "1" else []

    return templates.TemplateResponse(
        "employees.html",
        {
            "request": request,
            "employees": employees,
            "total_count": total_count,
            "page": page,
            "per_page": per_page,
            "per_page_selected": per_page,
            "total_pages": total_pages,
            "page_start": page_start,
            "prev_url": prev_url,
            "next_url": next_url,
            "role_order": ROLE_ORDER,
            "active_page": "employees",
            "query": q or "",
            "filter_role": role or "",
            "filter_category": category or "",
            "filter_cli_status": cli_status or "",
            "filter_gradation": gradation or "",
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
            "employee_return_to": employee_return_to,
            "employee_return_to_query": employee_return_to_query,
            "sync_notice": sync_notice or "",
            "sync_warning": sync_warning or "",
            "sync_error": sync_error or "",
            "sync_backup": request.query_params.get("sync_backup", ""),
            "sync_backup_label": _latest_employee_sync_backup()[1],
            "google_sync_ready": _google_sheet_sync_ready(),
            "google_sync_range": ", ".join(GOOGLE_EMPLOYEE_STATION_TABS),
            "retired_cleanup_notice": retired_cleanup_notice or "",
            "retired_cleanup_error": retired_cleanup_error or "",
            "retired_preview_rows": retired_preview_rows,
            "retired_preview_count": len(retired_preview_rows),
        },
    )


def _run_google_sheet_sync(session: Session, *, commit_changes: bool) -> dict[str, object]:
    backup_label = ""
    if commit_changes:
        try:
            backup_label = _create_employee_sync_backup()
        except Exception:
            backup_label = ""
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
        added_rows, updated_rows = _import_employee_rows(
            session,
            rows,
            source_label=f"Google Sheet ({sheet_name})",
            working_at_override=working_at,
            warnings=warnings,
            sync_details=sync_details,
            sync_stats=sync_stats,
            global_pf_counts=global_pf_counts,
            global_hrms_counts=global_hrms_counts,
            commit_changes=commit_changes,
        )
        added += added_rows
        updated += updated_rows

    if not commit_changes:
        session.rollback()

    unchanged = sync_stats.get("unchanged", 0)
    skipped = sync_stats.get("skipped", 0)
    if added == 0 and updated == 0 and skipped == 0:
        message_text = "No change found"
    else:
        prefix = "Google Sheet sync complete" if commit_changes else "Preview ready"
        message_text = f"{prefix}: {added} added, {updated} updated, {unchanged} unchanged, {skipped} skipped from {source_label}."
    warning_text = ""
    if warnings:
        preview = "; ".join(warnings[:3])
        if len(warnings) > 3:
            preview += f"; and {len(warnings) - 3} more"
        warning_text = f"Auto-corrected {len(warnings)} date value(s): {preview}"
    backup_notice = ""
    if commit_changes:
        backup_notice = f"Backup saved: {backup_label}" if backup_label else "Backup failed to save."
    added_details = [item for item in sync_details if item.startswith("Added ")]
    updated_details = [item for item in sync_details if item.startswith("Updated ")]
    auto_corrected_details = [item for item in sync_details if item.startswith("Auto-corrected ")]
    return {
        "message": message_text,
        "warning_message": warning_text,
        "warning_details": warnings,
        "sync_details": sync_details,
        "added_details": added_details,
        "updated_details": updated_details,
        "auto_corrected_details": auto_corrected_details,
        "deleted_details": [],
        "backup_notice": backup_notice,
        "has_changes": bool(added or updated or skipped),
    }


@app.post("/employees/sync-google-preview")
def preview_google_sheet_sync(
    request: Request,
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    wants_json = request.headers.get("x-requested-with", "").lower() == "fetch"
    try:
        _validate_sensitive_action_password(action_password)
        payload = _run_google_sheet_sync(session, commit_changes=False)
        if wants_json:
            return JSONResponse({"ok": True, **payload, "preview": True})
        return RedirectResponse(
            url=f"/employees?sync_notice={urlparse.quote(str(payload['message']))}#google-sync-card",
            status_code=303,
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Google Sheet preview failed."
        if wants_json:
            return JSONResponse({"ok": False, "message": detail}, status_code=exc.status_code)
        return RedirectResponse(url=f"/employees?sync_error={urlparse.quote(detail)}#google-sync-card", status_code=303)
    except Exception as exc:
        if wants_json:
            return JSONResponse({"ok": False, "message": str(exc)}, status_code=500)
        return RedirectResponse(url=f"/employees?sync_error={urlparse.quote(str(exc))}#google-sync-card", status_code=303)


@app.post("/employees/sync-google")
def sync_employees_from_google_sheet(
    request: Request,
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    wants_json = request.headers.get("x-requested-with", "").lower() == "fetch"
    try:
        _validate_sensitive_action_password(action_password)
        payload = _run_google_sheet_sync(session, commit_changes=True)
        if wants_json:
            return JSONResponse({"ok": True, **payload})
        message = urlparse.quote(str(payload["message"]))
        redirect_url = f"/employees?sync_notice={message}&sync_backup={urlparse.quote(str(payload['backup_notice']))}#google-sync-card"
        if payload.get("warning_message"):
            redirect_url = (
                f"/employees?sync_notice={message}&sync_warning={urlparse.quote(str(payload['warning_message']))}"
                f"&sync_backup={urlparse.quote(str(payload['backup_notice']))}#google-sync-card"
            )
        return RedirectResponse(url=redirect_url, status_code=303)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Google Sheet sync failed."
        if wants_json:
            return JSONResponse({"ok": False, "message": detail}, status_code=exc.status_code)
        return RedirectResponse(url=f"/employees?sync_error={urlparse.quote(detail)}#google-sync-card", status_code=303)
    except Exception as exc:
        if wants_json:
            return JSONResponse({"ok": False, "message": str(exc)}, status_code=500)
        return RedirectResponse(url=f"/employees?sync_error={urlparse.quote(str(exc))}#google-sync-card", status_code=303)


@app.post("/employees/restore-backup")
def restore_employee_backup(
    request: Request,
    action_password: str = Form(...),
):
    _validate_sensitive_action_password(action_password)
    backup_path, backup_label = _latest_employee_sync_backup()
    if not backup_path:
        raise HTTPException(status_code=400, detail="No backup available yet.")
    engine.dispose()
    shutil.copy2(backup_path, DB_PATH)
    notice = urlparse.quote(f"Backup restored: {backup_label}")
    return RedirectResponse(url=f"/employees?sync_notice={notice}#google-sync-card", status_code=303)


@app.post("/employees/retired-cleanup-preview")
def preview_retired_employee_cleanup():
    return RedirectResponse(url="/employees?retired_preview=1#retired-cleanup-card", status_code=303)


@app.post("/employees/retired-cleanup-apply")
def apply_retired_employee_cleanup(
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        retired_rows = _retired_employee_candidates(session)
        if not retired_rows:
            notice = urlparse.quote("No retired employees found for deletion.")
            return RedirectResponse(url=f"/employees?retired_cleanup_notice={notice}#retired-cleanup-card", status_code=303)
        deleted_count = len(retired_rows)
        for employee in retired_rows:
            session.delete(employee)
        session.commit()
        notice = urlparse.quote(f"Deleted {deleted_count} employee(s) whose retirement date is before {date.today().isoformat()}.")
        return RedirectResponse(url=f"/employees?retired_cleanup_notice={notice}#retired-cleanup-card", status_code=303)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Retired employee delete failed."
        return RedirectResponse(url=f"/employees?retired_cleanup_error={urlparse.quote(detail)}#retired-cleanup-card", status_code=303)
    except Exception as exc:
        return RedirectResponse(url=f"/employees?retired_cleanup_error={urlparse.quote(str(exc))}#retired-cleanup-card", status_code=303)


@app.get("/employees/{emp_id}")
def edit_employee_page(emp_id: int, request: Request, session: Session = Depends(get_session)):
    _sync_retired_employees(session, date.today())
    employee = session.get(Employee, emp_id)
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    return_to = _safe_return_to(
        request.query_params.get("return_to") or request.headers.get("referer"),
        fallback="/employees",
    )
    return templates.TemplateResponse(
        "employees_edit.html",
        {
            "request": request,
            "employee": employee,
            "role_order": ROLE_ORDER,
            "cli_choices": _cli_choice_rows(session),
            "active_page": "employees",
            "return_to": return_to,
        },
    )


@app.post("/employees/{emp_id}")
def update_employee(
    emp_id: int,
    return_to: Optional[str] = Form(None),
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
    cli_id: Optional[str] = Form(None),
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

    retirement_value = to_date(retirement_date)
    if retirement_value and retirement_value <= date.today():
        session.delete(employee)
        session.commit()
        return RedirectResponse("/employees", status_code=303)

    employee.name = name.strip()
    employee.role = role.strip()
    employee.retirement_date = retirement_value
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
    employee.cli, employee.cli_id = _canonicalize_cli_name(cli, cli_id)
    employee.pme_due = to_date(pme_due)
    employee.technical_due = to_date(technical_due)
    employee.transportation_due = to_date(transportation_due)

    session.add(employee)
    session.commit()
    return RedirectResponse(_safe_return_to(return_to, fallback="/employees"), status_code=303)


@app.post("/employees/{emp_id}/delete")
def delete_employee(
    emp_id: int,
    return_to: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    employee = session.get(Employee, emp_id)
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    session.delete(employee)
    session.commit()
    return RedirectResponse(_safe_return_to(return_to, fallback="/employees"), status_code=303)


@app.get("/uploads")
def uploads_page(request: Request):
    return templates.TemplateResponse("uploads.html", _uploads_context(request))


def _uploads_context(
    request: Request,
    *,
    update_error: str = "",
    update_notice: str = "",
    update_warning: str = "",
    update_details: Optional[list[str]] = None,
    warning_details: Optional[list[str]] = None,
    update_mismatch_actions: Optional[list[dict[str, object]]] = None,
) -> dict[str, object]:
    saved_review_report = _load_employee_master_review_report()
    review_reports = list(saved_review_report.get("history") or [])
    latest_review = review_reports[0] if review_reports else saved_review_report
    mismatch_actions = update_mismatch_actions
    if mismatch_actions is None:
        mismatch_actions = list(latest_review.get("update_mismatch_actions") or _load_employee_master_mismatch_actions())
    if not update_notice:
        update_notice = str(latest_review.get("update_notice") or "")
    if not update_warning:
        update_warning = str(latest_review.get("update_warning") or "")
    if update_details is None:
        update_details = list(latest_review.get("update_details") or [])
    if warning_details is None:
        warning_details = list(latest_review.get("warning_details") or [])
    service_snapshot_ready = EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE.exists()
    service_snapshot_saved_at = (
        datetime.fromtimestamp(EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
        if service_snapshot_ready
        else ""
    )
    source_snapshot_ready = EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.exists()
    source_snapshot_saved_at = (
        datetime.fromtimestamp(EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
        if source_snapshot_ready
        else ""
    )
    return {
        "request": request,
        "active_page": "uploads",
        "role_order": ROLE_ORDER,
        "update_error": update_error,
        "update_notice": update_notice,
        "update_warning": update_warning,
        "update_details": update_details or [],
        "warning_details": warning_details or [],
        "update_mismatch_actions": mismatch_actions or [],
        "update_preview_ready": False,
        "update_preview_password": "",
        "update_added_details": list(latest_review.get("update_added_details") or []),
        "update_updated_details": list(latest_review.get("update_updated_details") or []),
        "update_deduplicated_details": list(latest_review.get("update_deduplicated_details") or []),
        "review_report_saved_at": str(latest_review.get("saved_at") or ""),
        "review_reports": review_reports,
        "cleanup_error": "",
        "cleanup_notice": "",
        "cleanup_summary": None,
        "cleanup_groups": [],
        "cleanup_details": [],
        "cleanup_saved_at": "",
        "source_snapshot_ready": source_snapshot_ready,
        "source_snapshot_saved_at": source_snapshot_saved_at,
        "service_snapshot_ready": service_snapshot_ready,
        "service_snapshot_saved_at": service_snapshot_saved_at,
    }


def _uploads_template_response(
    request: Request,
    *,
    status_code: int = 200,
    update_error: str = "",
    update_notice: str = "",
    update_warning: str = "",
    update_details: Optional[list[str]] = None,
    warning_details: Optional[list[str]] = None,
    update_mismatch_actions: Optional[list[dict[str, object]]] = None,
):
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            update_error=update_error,
            update_notice=update_notice,
            update_warning=update_warning,
            update_details=update_details,
            warning_details=warning_details,
            update_mismatch_actions=update_mismatch_actions,
        ),
        status_code=status_code,
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
    _sync_retired_employees(session, today)
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
    today = date.today()
    _sync_retired_employees(session, today)
    _sync_employee_duplicates(session)
    start = _parse_date_cookie(request, "reports_start_date", start_date)
    end = _parse_date_cookie(request, "reports_end_date", end_date)
    if end < start:
        start, end = end, start
    horizon_months = 0
    employees = fetch_active_employees(session, today)
    cli_distribution = build_cli_distribution(employees)
    role_headers, working_summary = build_working_location_summary(employees)
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
    roster_cli_id: str | None = None,
    roster_role: str | None = None,
    roster_gradation: str | None = None,
    roster_cli_status: str | None = None,
    grading_update_error: str = "",
    grading_update_notice: str = "",
    grading_update_warning: str = "",
    grading_update_details: list[str] | None = None,
    grading_warning_details: list[str] | None = None,
    nomination_mismatch_actions: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    init_db()
    today = date.today()
    _sync_retired_employees(session, today)
    _sync_employee_duplicates(session)
    employees = [employee for employee in fetch_active_employees(session, today) if not _is_hidden_employee_role(employee.role)]
    grading_meta = _load_li_grading_metadata()
    grading_report_date = coerce_report_date(grading_meta.get("report_date"))
    saved_at_raw = grading_meta.get("saved_at", "")
    grading_saved_at = ""
    if saved_at_raw:
        try:
            grading_saved_at = datetime.fromisoformat(saved_at_raw).strftime("%d/%m/%Y %I:%M %p")
        except ValueError:
            grading_saved_at = saved_at_raw
    cli_bio_reference_rows = _load_cli_bio_reference_rows(session)
    latest_cli_bio_source = str(next((row.get("source_file") for row in cli_bio_reference_rows if row.get("source_file")), "") or "")
    latest_cli_bio_saved_at = str(next((row.get("updated_at") for row in cli_bio_reference_rows if row.get("updated_at")), "") or "")
    cli_opts = sorted(
        {
            str(row.get("cli_name") or "").strip()
            for row in cli_bio_reference_rows
            if str(row.get("cli_name") or "").strip()
        }
        or {(e.cli or "").strip() for e in employees if (e.cli or "").strip()}
    )
    role_opts = ROLE_ORDER + sorted({e.role for e in employees if e.role not in ROLE_ORDER})
    gradation_opts = sorted({e.gradation for e in employees if e.gradation})
    cli_distribution = build_cli_distribution(employees, cli_bio_reference_rows)
    for row in cli_distribution:
        detail_params = [f"roster_cli={urlparse.quote(str(row.get('cli') or ''))}"]
        if row.get("cli_id"):
            detail_params.append(f"roster_cli_id={urlparse.quote(str(row.get('cli_id') or ''))}")
        row["detail_href"] = f"/cli?{'&'.join(detail_params)}#cli-roster"
        row_cli_name, row_cli_id = _canonicalize_cli_name(row.get("cli"), row.get("cli_id"))
        selected_cli_name, selected_cli_id = _canonicalize_cli_name(roster_cli, roster_cli_id)
        row["selected"] = bool(
            (selected_cli_id and row_cli_id == selected_cli_id)
            or (selected_cli_name and _cli_names_equivalent(row_cli_name, selected_cli_name))
        )

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
        selected_cli_name, selected_cli_id = _canonicalize_cli_name(roster_cli, roster_cli_id)
        cli_roster = [
            e
            for e in cli_roster
            if (
                lambda employee_cli_name, employee_cli_id: (
                    (selected_cli_id and employee_cli_id == selected_cli_id)
                    or (selected_cli_name and _cli_names_equivalent(employee_cli_name, selected_cli_name))
                    or (not selected_cli_id and not selected_cli_name and roster_cli.strip().lower() in _employee_cli_label(e).lower())
                )
            )(*_canonicalize_cli_name(e.cli, e.cli_id))
        ]
    if roster_role:
        cli_roster = [e for e in cli_roster if e.role == roster_role]
    if roster_gradation:
        grad_lower = roster_gradation.lower()
        cli_roster = [e for e in cli_roster if e.gradation and grad_lower in e.gradation.lower()]
    cli_roster = sorted(cli_roster, key=lambda e: ((e.cli or "").strip().lower(), role_sort_key(e.role), e.name))
    if nomination_mismatch_actions is None:
        nomination_mismatch_actions = _load_cli_nomination_mismatch_actions()

    return {
        "request": request,
        "active_page": "cli",
        "cli_distribution": cli_distribution,
        "cli_distribution_totals_all": totals_all,
        "cli_distribution_breakdown": [],
        "cli_distribution_totals": {"A": 0, "B": 0, "C": 0, "total": 0},
        "cli_distribution_detail_label": roster_cli or "",
        "cli_bio_reference_rows": cli_bio_reference_rows,
        "latest_cli_bio_source": latest_cli_bio_source,
        "latest_cli_bio_saved_at": latest_cli_bio_saved_at,
        "cli_roster": cli_roster,
        "cli_opts": cli_opts,
        "role_opts": role_opts,
        "gradation_opts": gradation_opts,
        "roster_name": roster_name or "",
        "roster_cli": roster_cli or "",
        "roster_cli_id": roster_cli_id or "",
        "roster_role": roster_role or "",
        "roster_gradation": roster_gradation or "",
        "roster_cli_status": roster_cli_status or "",
        "roster_open": True,
        "grading_update_error": grading_update_error,
        "grading_update_notice": grading_update_notice,
        "grading_update_warning": grading_update_warning,
        "grading_update_details": grading_update_details or [],
        "grading_warning_details": grading_warning_details or [],
        "nomination_mismatch_actions": nomination_mismatch_actions or [],
        "grading_source_name": grading_meta.get("filename", ""),
        "grading_report_date": grading_report_date.strftime("%d-%m-%Y") if grading_report_date else "",
        "grading_saved_at": grading_saved_at,
    }


@app.get("/cli")
def cli_page(
    request: Request,
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_cli_id: Optional[str] = None,
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
            roster_cli_id=roster_cli_id,
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
    return templates.TemplateResponse("top_performer.html", _top_performer_context(request, **_load_top_performer_store()))


@app.post("/top-performer")
async def generate_top_performer(
    request: Request,
    files: list[UploadFile] = File(...),
    minimum_runs: int = Form(3),
):
    valid_files = [upload for upload in files if (upload.filename or "").strip()]
    if not valid_files:
        return templates.TemplateResponse(
            "top_performer.html",
            _top_performer_context(request, warnings=["Please upload at least one ranking file."]),
        )
    results: list[dict[str, object]] = []
    warnings: list[str] = []
    try:
        for upload in valid_files:
            rows, filename = _top_performer_read_tabular_file(upload)
            result, file_warnings = _build_top_performer_result(rows, filename, max(0, minimum_runs))
            results.append(result)
            warnings.extend(file_warnings)
    except Exception as exc:
        return templates.TemplateResponse(
            "top_performer.html",
            _top_performer_context(request, minimum_runs=max(0, minimum_runs), warnings=[f"Top performer generation failed: {exc}"]),
        )
    payload = _top_performer_blank_state()
    payload.update(
        {
            "saved_at_label": datetime.now(IST).strftime("%d-%m-%Y %I:%M %p"),
            "warnings": warnings,
            "minimum_runs": max(0, minimum_runs),
            "summary": {
                "file_count": len(results),
                "overall_rows": sum(int(result.get("row_count") or 0) for result in results),
                "overall_eligible": sum(int(result.get("eligible_count") or 0) for result in results),
            },
            "results": results,
            "comparison": None,
        }
    )
    _save_top_performer_store(payload)
    return templates.TemplateResponse("top_performer.html", _top_performer_context(request, **payload))


@app.post("/top-performer/reset")
def reset_top_performer(request: Request):
    store = _load_top_performer_store()
    store["summary"] = None
    store["results"] = []
    store["warnings"] = []
    _save_top_performer_store(store)
    return RedirectResponse(url="/top-performer", status_code=303)


@app.post("/top-performer/clear-stored")
def clear_top_performer_stored(
    action_password: str = Form(...),
):
    _validate_sensitive_action_password(action_password)
    store = _top_performer_blank_state()
    _save_top_performer_store(store)
    if TOP_PERFORMER_PHOTO_DIR.exists():
        for path in TOP_PERFORMER_PHOTO_DIR.iterdir():
            if path.is_file():
                path.unlink(missing_ok=True)
    return RedirectResponse(url="/top-performer", status_code=303)


@app.post("/top-performer/photo")
async def save_top_performer_photo(
    crew_name: str = Form(...),
    photo_file: UploadFile = File(...),
):
    crew_name = crew_name.strip()
    if not crew_name:
        raise HTTPException(status_code=400, detail="Crew name is required.")
    extension = Path(photo_file.filename or "").suffix.lower()
    if extension not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise HTTPException(status_code=400, detail="Upload a PNG, JPG, JPEG, or WEBP photo.")
    _ensure_top_performer_dirs()
    target_name = f"{_top_performer_slug(crew_name)}{extension}"
    target_path = TOP_PERFORMER_PHOTO_DIR / target_name
    target_path.write_bytes(await photo_file.read())
    photo_url = f"/static/top_performer_photos/{target_name}"
    store = _load_top_performer_store()
    for result in store.get("results", []):
        if not isinstance(result, dict):
            continue
        for row in result.get("top_rows", []):
            if isinstance(row, dict) and str(row.get("crew_name") or "").strip().casefold() == crew_name.casefold():
                row["photo_url"] = photo_url
    comparison = store.get("comparison")
    if isinstance(comparison, dict):
        for row in comparison.get("rows", []):
            if isinstance(row, dict) and str(row.get("crew_name") or "").strip().casefold() == crew_name.casefold():
                row["photo_url"] = photo_url
    _save_top_performer_store(store)
    return RedirectResponse(url="/top-performer", status_code=303)


@app.post("/top-performer/compare")
async def compare_top_performer_months(
    request: Request,
    previous_file: UploadFile = File(...),
    current_file: UploadFile = File(...),
    minimum_runs: int = Form(3),
):
    try:
        previous_rows, previous_filename = _top_performer_read_tabular_file(previous_file)
        current_rows, current_filename = _top_performer_read_tabular_file(current_file)
        comparison, warnings = _build_top_performer_comparison(
            previous_rows,
            previous_filename,
            current_rows,
            current_filename,
            max(0, minimum_runs),
        )
    except Exception as exc:
        store = _load_top_performer_store()
        store["minimum_runs"] = max(0, minimum_runs)
        store["warnings"] = [f"Monthly comparison failed: {exc}"]
        _save_top_performer_store(store)
        return templates.TemplateResponse("top_performer.html", _top_performer_context(request, **store))
    store = _load_top_performer_store()
    store["minimum_runs"] = max(0, minimum_runs)
    store["warnings"] = warnings
    store["comparison"] = comparison
    if not store.get("saved_at_label"):
        store["saved_at_label"] = datetime.now(IST).strftime("%d-%m-%Y %I:%M %p")
    _save_top_performer_store(store)
    return templates.TemplateResponse("top_performer.html", _top_performer_context(request, **store))


@app.post("/top-performer/comparison/reset")
def reset_top_performer_comparison():
    store = _load_top_performer_store()
    store["comparison"] = None
    store["warnings"] = []
    _save_top_performer_store(store)
    return RedirectResponse(url="/top-performer", status_code=303)


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


def _build_ssts_report_response(
    request: Request,
    session: Session,
    *,
    force: bool = False,
    report_tab: str = "online_offline",
    selected_day: str | None = None,
    analysis_day: str | None = None,
    selected_analysis_rake: str | None = None,
    detail_view: str | None = None,
    pf_day: str | None = None,
    pf_speed_threshold: str | None = None,
    pf_task_id: str | None = None,
    pf_train: str | None = None,
    pf_detail_mode: str | None = None,
    junk_cleanup_notice: str = "",
    junk_cleanup_error: str = "",
    status_code: int = 200,
):
    sync_result = refresh_ssts_snapshot(session, force=force)
    active_report_tab = report_tab if report_tab in {"online_offline", "pf_entering"} else "online_offline"
    selected_day_value = _parse_report_date(selected_day)
    analysis_day_value = _parse_report_date(analysis_day)
    selected_analysis_rake_value: int | None = None
    if selected_analysis_rake:
        try:
            selected_analysis_rake_value = int(selected_analysis_rake)
        except ValueError:
            selected_analysis_rake_value = None
    pf_day_value = selected_day_value or date.today()
    pf_speed_threshold_value = _parse_pf_speed_threshold(pf_speed_threshold)
    pf_speed_options = [{"value": value, "label": str(value)} for value in range(40, 51)] + [
        {"value": 51, "label": "Above 50"}
    ]
    parsed_pf_day = _parse_report_date(pf_day)
    if parsed_pf_day is not None:
        pf_day_value = parsed_pf_day
    context = build_ssts_report_context(
        session,
        selected_day=selected_day_value,
        analysis_day=analysis_day_value,
        selected_analysis_rake=selected_analysis_rake_value,
    )
    pf_context = {
        "pf_report_day": pf_day_value.isoformat(),
        "pf_report_day_label": pf_day_value.strftime("%d-%m-%Y"),
        "pf_report_rows": [],
        "pf_speed_threshold": pf_speed_threshold_value,
        "pf_speed_threshold_label": _pf_speed_threshold_label(pf_speed_threshold_value),
        "pf_speed_options": pf_speed_options,
        "pf_daily_report_rows": [],
        "pf_detailed_daily_report_rows": [],
        "pf_detailed_daily_spike_count": 0,
        "pf_detailed_daily_spike_rows": [],
        "pf_analysis_summary_rows": [],
        "pf_analysis_selected_rows": [],
        "pf_analysis_selected_train": "",
        "pf_analysis_selected_mode": "raw",
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
            if task_payload.get("speed_threshold") is not None:
                resolved_threshold = _parse_pf_speed_threshold(task_payload.get("speed_threshold"))
                pf_context["pf_speed_threshold"] = resolved_threshold
                pf_context["pf_speed_threshold_label"] = _pf_speed_threshold_label(resolved_threshold)
            if task_payload.get("status") == "completed":
                result = task_payload.get("result")
                if isinstance(result, dict):
                    result = _pf_hydrate_chart_links_in_result(result)
                    pf_context.update(
                        {
                            key: value
                            for key, value in result.items()
                            if key not in {"pf_analysis_detail_rows_by_train", "pf_detailed_detail_rows_by_train"}
                        }
                    )
                    raw_detail_rows_by_train = result.get("pf_analysis_detail_rows_by_train")
                    clean_detail_rows_by_train = result.get("pf_detailed_detail_rows_by_train")
                    selected_mode = "clean" if pf_detail_mode == "clean" else "raw"
                    selected_detail_rows_by_train = (
                        clean_detail_rows_by_train if selected_mode == "clean" else raw_detail_rows_by_train
                    )
                    if isinstance(selected_detail_rows_by_train, dict):
                        selected_train_value = pf_train or (
                            str(pf_context["pf_analysis_summary_rows"][0].get("train_no") or "")
                            if pf_context["pf_analysis_summary_rows"]
                            else ""
                        )
                        pf_context["pf_analysis_selected_train"] = selected_train_value
                        pf_context["pf_analysis_selected_mode"] = selected_mode
                        selected_rows = selected_detail_rows_by_train.get(selected_train_value, [])
                        pf_context["pf_analysis_selected_rows"] = selected_rows if isinstance(selected_rows, list) else []
            elif task_payload.get("status") == "error":
                pf_context["pf_report_error"] = str(task_payload.get("message") or "PF analysis failed.")
        elif pf_task_id and parsed_pf_day is not None:
            try:
                rebuilt_result = _build_ssts_pf_speed_analysis_result(pf_day_value, pf_speed_threshold_value)
                rebuilt_result = _pf_hydrate_chart_links_in_result(rebuilt_result)
                pf_context.update(rebuilt_result)
                pf_context["pf_analysis_status"] = "completed"
                pf_context["pf_analysis_message"] = "Analysis restored after status refresh."
                pf_context["pf_analysis_task_id"] = pf_task_id or ""
                selected_mode = "clean" if pf_detail_mode == "clean" else "raw"
                selected_detail_rows_by_train = (
                    rebuilt_result.get("pf_detailed_detail_rows_by_train")
                    if selected_mode == "clean"
                    else rebuilt_result.get("pf_analysis_detail_rows_by_train")
                )
                if isinstance(selected_detail_rows_by_train, dict):
                    selected_train_value = pf_train or (
                        str(pf_context["pf_analysis_summary_rows"][0].get("train_no") or "")
                        if pf_context["pf_analysis_summary_rows"]
                        else ""
                    )
                    pf_context["pf_analysis_selected_train"] = selected_train_value
                    pf_context["pf_analysis_selected_mode"] = selected_mode
                    selected_rows = selected_detail_rows_by_train.get(selected_train_value, [])
                    pf_context["pf_analysis_selected_rows"] = selected_rows if isinstance(selected_rows, list) else []
            except Exception as exc:
                pf_context["pf_report_error"] = f"PF analysis restore failed: {exc}"
    latest_run = context.get("latest_run")
    junk_cleanup_summary = _ssts_cleanup_junk_summary()
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
            "ssts_cleanup_summary": sync_result.get("cleanup_summary") or {"deleted_runs": 0, "deleted_snapshots": 0},
            "ssts_junk_cleanup_summary": junk_cleanup_summary,
            "ssts_junk_cleanup_notice": junk_cleanup_notice,
            "ssts_junk_cleanup_error": junk_cleanup_error,
            "ssts_retention_days": SSTS_SNAPSHOT_RETENTION_DAYS,
            "IST": IST,
            "active_detail_view": detail_view if detail_view in {"recent_offline", "recently_online"} else None,
            **context,
            **pf_context,
        },
        status_code=status_code,
    )


@app.get("/ssts-report")
def ssts_report_page(
    request: Request,
    force: int = 0,
    report_tab: str = "online_offline",
    selected_day: str | None = None,
    analysis_day: str | None = None,
    selected_analysis_rake: str | None = None,
    detail_view: str | None = None,
    pf_day: str | None = None,
    pf_speed_threshold: str | None = None,
    pf_task_id: str | None = None,
    pf_train: str | None = None,
    pf_detail_mode: str | None = None,
    session: Session = Depends(get_session),
):
    return _build_ssts_report_response(
        request,
        session,
        force=bool(force),
        report_tab=report_tab,
        selected_day=selected_day,
        analysis_day=analysis_day,
        selected_analysis_rake=selected_analysis_rake,
        detail_view=detail_view,
        pf_day=pf_day,
        pf_speed_threshold=pf_speed_threshold,
        pf_task_id=pf_task_id,
        pf_train=pf_train,
        pf_detail_mode=pf_detail_mode,
    )


@app.post("/ssts-report/cleanup-junk")
def ssts_report_cleanup_junk(
    request: Request,
    report_tab: str = Form("online_offline"),
    selected_day: str = Form(""),
    analysis_day: str = Form(""),
    selected_analysis_rake: str = Form(""),
    detail_view: str = Form(""),
    pf_day: str = Form(""),
    pf_speed_threshold: str = Form(""),
    pf_task_id: str = Form(""),
    pf_train: str = Form(""),
    pf_detail_mode: str = Form(""),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        cleanup_result = _delete_ssts_junk_files()
        cleanup_notice = (
            f"Deleted {cleanup_result['deleted_count']} junk item(s), freed {cleanup_result['freed_label']}."
            if cleanup_result["deleted_count"]
            else "No removable junk files were found."
        )
        return _build_ssts_report_response(
            request,
            session,
            report_tab=report_tab,
            selected_day=selected_day or None,
            analysis_day=analysis_day or None,
            selected_analysis_rake=selected_analysis_rake or None,
            detail_view=detail_view or None,
            pf_day=pf_day or None,
            pf_speed_threshold=pf_speed_threshold or None,
            pf_task_id=pf_task_id or None,
            pf_train=pf_train or None,
            pf_detail_mode=pf_detail_mode or None,
            junk_cleanup_notice=cleanup_notice,
        )
    except HTTPException as exc:
        return _build_ssts_report_response(
            request,
            session,
            report_tab=report_tab,
            selected_day=selected_day or None,
            analysis_day=analysis_day or None,
            selected_analysis_rake=selected_analysis_rake or None,
            detail_view=detail_view or None,
            pf_day=pf_day or None,
            pf_speed_threshold=pf_speed_threshold or None,
            pf_task_id=pf_task_id or None,
            pf_train=pf_train or None,
            pf_detail_mode=pf_detail_mode or None,
            junk_cleanup_error=exc.detail if isinstance(exc.detail, str) else "Cleanup failed.",
            status_code=exc.status_code,
        )


@app.get("/ssts-report/pf-chart")
def ssts_pf_chart_page(
    request: Request,
    train_date: str,
    train_no: str,
    device_id: str = "",
    org: str = "",
    dep: str = "",
    dest: str = "",
    arr: str = "",
    station: str = "",
    start_pos: str = "",
    end_pos: str = "",
):
    try:
        report_day = date.fromisoformat(train_date)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid train_date.")

    token = ""
    source = {
        "train_date_iso": report_day.isoformat(),
        "train_no": train_no,
        "device_id": _coerce_int(device_id) or device_id,
        "org": org,
        "train_dep_raw": dep,
        "dest": dest,
        "train_arr_raw": arr,
    }
    chart_points: list[dict[str, object]] = []
    train_rows: list[dict[str, object]] = []
    error_message = ""
    try:
        token = fetch_ssts_token()
        train_report_rows = fetch_ssts_trains_report(report_day, token)
        train_match = next(
            (
                item for item in train_report_rows
                if str(item.get("train_no") or "").strip() == train_no.strip()
            ),
            None,
        )
        if isinstance(train_match, dict):
            actual_dep = train_match.get("act_dep") or train_match.get("dep")
            actual_arr = train_match.get("act_arr") or train_match.get("arr")
            if actual_dep not in (None, ""):
                source["train_dep_raw"] = actual_dep
            if actual_arr not in (None, ""):
                source["train_arr_raw"] = actual_arr
        chart_points = _fetch_ssts_positions(source, token)
        train_rows = _build_pf_report_rows_for_train(
            train_match if isinstance(train_match, dict) else {
                "train_no": train_no,
                "device_id": _coerce_int(device_id) or device_id,
                "org": org,
                "dest": dest,
                "dep": source.get("train_dep_raw"),
                "arr": source.get("train_arr_raw"),
                "device_name": "",
                "crew_name": "",
            },
            report_day,
            token,
        )
    except (urlerror.URLError, RuntimeError, ValueError, json.JSONDecodeError) as exc:
        error_message = str(exc)

    categories = [_pf_chart_time_label(point) for point in chart_points]
    speed_series = [_pf_chart_speed_kmph(point) for point in chart_points]
    distance_series = _pf_chart_distance_series_km(chart_points)
    highlight_from = _coerce_int(start_pos)
    highlight_to = _coerce_int(end_pos)
    if highlight_from is not None and highlight_to is not None and highlight_to < highlight_from:
        highlight_from, highlight_to = highlight_to, highlight_from
    normalized_train_rows, highlight_from, highlight_to = _pf_normalize_station_windows_to_chart(
        train_rows,
        chart_points,
        len(chart_points),
        station,
        highlight_from,
        highlight_to,
    )
    station_plot_bands = _pf_build_station_plot_bands(normalized_train_rows, station, highlight_from, highlight_to)

    return templates.TemplateResponse(
        "ssts_pf_chart.html",
        {
            "request": request,
            "active_page": "ssts_report",
            "report_day": report_day.isoformat(),
            "report_day_label": report_day.strftime("%d-%m-%Y"),
            "train_no": train_no,
            "device_id": str(device_id or ""),
            "org": org,
            "dest": dest,
            "station": station,
            "chart_error": error_message,
            "chart_categories_json": json.dumps(categories),
            "chart_speed_json": json.dumps(speed_series),
            "chart_distance_json": json.dumps(distance_series),
            "station_plot_bands_json": json.dumps(station_plot_bands),
            "highlight_from": highlight_from,
            "highlight_to": highlight_to,
            "ssts_web_url": SSTS_WEB_URL,
        },
    )


def _queue_ssts_pf_analysis(report_day: date, speed_threshold: int) -> str:
    task_id = uuid4().hex
    _set_ssts_pf_analysis_task(
        task_id,
        status="pending",
        progress=2,
        message="Queued for analysis...",
        report_day=report_day.isoformat(),
        speed_threshold=speed_threshold,
        result=None,
    )
    worker = threading.Thread(target=_run_ssts_pf_analysis_task, args=(task_id, report_day, speed_threshold), daemon=True)
    worker.start()
    return task_id


@app.post("/ssts-report/pf-analysis/start")
async def start_ssts_pf_analysis(
    pf_day: str = Form(...),
    pf_speed_threshold: str = Form("40"),
):
    report_day = _parse_report_date(pf_day)
    if report_day is None:
        raise HTTPException(status_code=400, detail="Invalid PF analysis date.")
    speed_threshold = _parse_pf_speed_threshold(pf_speed_threshold)
    task_id = _queue_ssts_pf_analysis(report_day, speed_threshold)
    return JSONResponse(
        {
            "task_id": task_id,
            "status": "pending",
            "status_url": f"/ssts-report/pf-analysis/status?task_id={task_id}",
            "result_url": (
                f"/ssts-report?report_tab=pf_entering&pf_task_id={task_id}"
                f"&pf_day={report_day.isoformat()}&pf_speed_threshold={speed_threshold}"
            ),
        }
    )


@app.get("/ssts-report/pf-analysis/start")
def start_ssts_pf_analysis_fallback(
    pf_day: str,
    pf_speed_threshold: str = "40",
    anchor: str | None = None,
):
    report_day = _parse_report_date(pf_day)
    if report_day is None:
        raise HTTPException(status_code=400, detail="Invalid PF analysis date.")
    speed_threshold = _parse_pf_speed_threshold(pf_speed_threshold)
    task_id = _queue_ssts_pf_analysis(report_day, speed_threshold)
    safe_anchor = ""
    if anchor in {"pf-daily-report", "pf-detailed-daily-report"}:
        safe_anchor = f"#{anchor}"
    return RedirectResponse(
        url=(
            f"/ssts-report?report_tab=pf_entering&pf_task_id={task_id}"
            f"&pf_day={report_day.isoformat()}&pf_speed_threshold={speed_threshold}{safe_anchor}"
        ),
        status_code=303,
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
            "result_url": (
                f"/ssts-report?report_tab=pf_entering&pf_task_id={task_id}"
                f"&pf_day={task_payload.get('report_day') or ''}"
                f"&pf_speed_threshold={_parse_pf_speed_threshold(task_payload.get('speed_threshold'))}"
            ),
        }
    )


@app.get("/reports/cli-distribution.xlsx")
def download_cli_distribution(session: Session = Depends(get_session)):
    today = date.today()
    _sync_retired_employees(session, today)
    employees = fetch_active_employees(session, today)
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
    title, headers, rows, report_date_label, cell_classes = _coerce_export_table_payload(payload)

    wb = Workbook()
    ws = wb.active
    ws.title = _sanitize_excel_sheet_title(title)
    column_count = max(1, len(headers))
    if report_date_label:
        ws.append([title])
        ws.append([report_date_label])
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

    alert_font = Font(bold=True, color="C41A1A")
    crew_alert_font = Font(bold=True, color="FF9F43")
    for row_offset, class_row in enumerate(cell_classes, start=1):
        sheet_row = header_row_index + row_offset
        for column_index, class_text in enumerate(class_row, start=1):
            class_name = str(class_text or "").lower()
            if "station-alert" in class_name or "pf-speed-alert" in class_name:
                ws.cell(row=sheet_row, column=column_index).font = alert_font
            if "crew-alert" in class_name:
                ws.cell(row=sheet_row, column=column_index).font = crew_alert_font

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

    title, headers, rows, report_date_label, cell_classes = _coerce_export_table_payload(payload)
    pdf_bytes = _build_table_pdf_bytes(title, headers, rows, report_date_label, cell_classes)
    filename = _build_pdf_export_filename(title, report_date_label)
    return StreamingResponse(
        BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{urlparse.quote(filename)}"},
    )


@app.post("/exports/table.pdf/form")
async def export_table_pdf_form(payload: str = Form(...)):
    try:
        parsed_payload = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid export payload.") from exc

    title, headers, rows, report_date_label, cell_classes = _coerce_export_table_payload(parsed_payload)
    pdf_bytes = _build_table_pdf_bytes(title, headers, rows, report_date_label, cell_classes)
    filename = _build_pdf_export_filename(title, report_date_label)
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

    today = date.today()
    _sync_retired_employees(session, today)
    role_norm = role.strip()
    retirement_value = to_date(retirement_date)
    existing = session.exec(
        select(Employee).where(Employee.name == name.strip(), Employee.role == role_norm)
    ).first()

    if retirement_value and retirement_value <= today:
        if existing:
            session.delete(existing)
            session.commit()
        return RedirectResponse("/employees", status_code=303)

    if existing:
        existing.hire_date = to_date(hire_date)
        existing.retirement_date = retirement_value
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
            retirement_date=retirement_value,
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
    _sync_retired_employees(session, plan_date)
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


def _normalize_employee_cli_names(session: Session) -> int:
    employees = session.exec(select(Employee)).all()
    canonical_by_id, alias_map, id_by_name = _build_cli_name_maps((employee.cli, employee.cli_id) for employee in employees)
    changed = 0
    for employee in employees:
        new_cli, new_cli_id = _canonicalize_cli_name(
            employee.cli,
            employee.cli_id,
            canonical_by_id=canonical_by_id,
            alias_map=alias_map,
            id_by_name=id_by_name,
        )
        if employee.cli != new_cli or employee.cli_id != new_cli_id:
            employee.cli = new_cli
            employee.cli_id = new_cli_id
            changed += 1
    if changed:
        session.commit()
    return changed


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
    "cliid": "cli_id",
    "cli_id": "cli_id",
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


GOOGLE_SYNC_UPDATED_FIELDS = {
    "CREW ID",
    "Designation",
    "PME Due",
    "Technical Due",
    "Transportation Due",
    "Gradation",
    "CLI",
    "Working At",
}


def _google_sync_change_label(changed_labels: list[str]) -> str:
    if any(label in GOOGLE_SYNC_UPDATED_FIELDS for label in changed_labels):
        return "Updated"
    return "Auto-corrected"


def _clean_import_text(value: object | None, *, blank_na: bool = False) -> str | None:
    text = str(value or "").strip()
    if blank_na and text.upper() in {"NA", "N/A"}:
        return None
    return text or None


def _normalize_import_name(value: object | None) -> str | None:
    text = _clean_import_text(value)
    if text is None:
        return None
    text = text.upper()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b", " ", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split()) or None


_NAME_TOKEN_EQUIVALENTS: dict[str, set[str]] = {
    "KR": {"KUMAR"},
    "KR.": {"KUMAR"},
    "KUMAR": {"KR"},
    "CH": {"CHANDRA"},
    "CH.": {"CHANDRA"},
    "CHANDRA": {"CH"},
}


def _names_almost_same(left: object | None, right: object | None) -> bool:
    left_name = _normalize_import_name(left)
    right_name = _normalize_import_name(right)
    if not left_name or not right_name:
        return False
    if left_name == right_name:
        return True
    left_tokens = left_name.split()
    right_tokens = right_name.split()
    if not left_tokens or not right_tokens:
        return False
    if left_tokens[-1] != right_tokens[-1]:
        return False
    shorter, longer = (left_tokens, right_tokens) if len(left_tokens) <= len(right_tokens) else (right_tokens, left_tokens)
    if len(longer) - len(shorter) > 1:
        return False
    i = 0
    j = 0
    while i < len(shorter) and j < len(longer):
        short_token = shorter[i]
        long_token = longer[j]
        if short_token == long_token:
            i += 1
            j += 1
            continue
        equivalent = _NAME_TOKEN_EQUIVALENTS.get(short_token, set())
        if long_token in equivalent:
            i += 1
            j += 1
            continue
        if len(short_token) == 1 and long_token.startswith(short_token):
            i += 1
            j += 1
            continue
        if len(long_token) == 1 and short_token.startswith(long_token):
            i += 1
            j += 1
            continue
        if i + 1 < len(shorter) and (short_token + shorter[i + 1]) == long_token:
            i += 2
            j += 1
            continue
        if j + 1 < len(longer) and short_token == (long_token + longer[j + 1]):
            i += 1
            j += 2
            continue
        return False
    return i == len(shorter) and j == len(longer)


def _emp_no_last5(value: object | None) -> str | None:
    text = _clean_import_text(value)
    if text is None:
        return None
    normalized = re.sub(r"[^A-Z0-9]", "", text.upper())
    if not normalized:
        return None
    return normalized[-5:] if len(normalized) >= 5 else normalized


def _emp_no_match_key(value: object | None) -> str | None:
    text = _clean_import_text(value)
    if text is None:
        return None
    normalized = re.sub(r"[^A-Z0-9]", "", text.upper())
    if not normalized:
        return None
    if len(normalized) > 1 and normalized.startswith("0"):
        normalized = normalized[1:]
    return normalized or None


def _emp_no_matches(left: object | None, right: object | None) -> bool:
    left_raw = _clean_import_text(left)
    right_raw = _clean_import_text(right)
    if left_raw is None or right_raw is None:
        return False
    if left_raw == right_raw:
        return True
    left_key = _emp_no_match_key(left_raw)
    right_key = _emp_no_match_key(right_raw)
    return left_key is not None and left_key == right_key


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
    commit_changes: bool = True,
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

    existing_cli_rows = session.exec(select(Employee.cli, Employee.cli_id)).all()
    canonical_by_id, alias_map, id_by_name = _build_cli_name_maps(existing_cli_rows)
    employees = session.exec(select(Employee)).all()
    prefer_existing_over_incoming = source_label.strip().lower().startswith("google sheet")

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

    def _find_pf_matches(value: str | None) -> list[Employee]:
        if not value:
            return []
        exact_matches = [employee for employee in employees if _clean_import_text(employee.pf_no) == value]
        if exact_matches:
            return exact_matches
        match_key = _emp_no_match_key(value)
        if match_key is None:
            return []
        return [employee for employee in employees if _emp_no_match_key(employee.pf_no) == match_key]

    def _prefer_target(existing_value: object, incoming_value: object, has_column: bool) -> object:
        if not has_column:
            return existing_value
        if not prefer_existing_over_incoming:
            return incoming_value
        if _employee_has_value(existing_value):
            return existing_value
        return incoming_value

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
                warnings.append(f"{source_label} {row_hint}: skipped because PF No {pf_no} appears multiple times in the Google Sheet.")
            continue
        if hrms and (source_hrms_counts.get(hrms, 0) > 1 or global_hrms_counts.get(hrms, 0) > 1):
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(f"{source_label} {row_hint}: skipped because HRMS {hrms} appears multiple times in the Google Sheet.")
            continue

        pf_matches = _find_pf_matches(pf_no)
        hrms_matches = session.exec(select(Employee).where(Employee.hrms == hrms)).all() if hrms else []
        if len(pf_matches) > 1:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(f"{source_label} {row_hint}: skipped because PF No {pf_no} matches multiple employees in the current database.")
            continue
        if len(hrms_matches) > 1:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(f"{source_label} {row_hint}: skipped because HRMS {hrms} matches multiple employees in the current database.")
            continue
        if pf_matches and hrms_matches and pf_matches[0].id != hrms_matches[0].id:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(f"{source_label} {row_hint}: skipped because PF No {pf_no} and HRMS {hrms} point to different employees.")
            continue

        both_matches = [employee for employee in pf_matches if hrms and employee.hrms == hrms] if pf_matches and hrms else []
        if len(both_matches) == 1:
            existing = both_matches[0]
        elif pf_matches:
            existing = pf_matches[0]
        elif hrms_matches:
            existing = hrms_matches[0]
        else:
            if crew_id:
                crew_matches = session.exec(select(Employee).where(Employee.crew_id == crew_id)).all()
                if len(crew_matches) > 1:
                    if sync_stats is not None:
                        sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
                    if warnings is not None:
                        warnings.append(f"{source_label} {row_hint}: skipped because CREW ID {crew_id} matches multiple employees in the current database.")
                    continue
                if len(crew_matches) == 1:
                    existing = crew_matches[0]

            if existing is None:
                exact_matches = session.exec(select(Employee).where(Employee.name == str(name).strip(), Employee.role == role)).all()
                if len(exact_matches) == 1 and not exact_matches[0].pf_no and not exact_matches[0].hrms:
                    existing = exact_matches[0]
                else:
                    normalized_name = _normalize_import_name(name)
                    role_candidates = session.exec(select(Employee).where(Employee.role == role)).all()
                    fallback_candidates = []
                    for candidate in role_candidates:
                        if not _names_almost_same(candidate.name, name):
                            continue
                        working_at_compatible = not working_at or not candidate.working_at or candidate.working_at == working_at
                        dob_compatible = dob is None or candidate.dob is None or candidate.dob == dob
                        if working_at_compatible and dob_compatible:
                            fallback_candidates.append(candidate)
                    if len(fallback_candidates) == 1:
                        existing = fallback_candidates[0]
                    elif len(fallback_candidates) > 1:
                        if sync_stats is not None:
                            sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
                        if warnings is not None:
                            warnings.append(f"{source_label} {row_hint}: skipped because name/role fallback matched multiple existing employees.")
                        continue

            if existing is None and pf_no is None and hrms is None and crew_id is None:
                if sync_stats is not None:
                    sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
                if warnings is not None:
                    warnings.append(f"{source_label} {row_hint}: skipped because PF No, HRMS, and CREW ID are blank and no unique existing employee match was found.")
                continue

        if hire_date is None and (has_col("hire_date") or has_col("doa") or has_col("dob") or has_col("retirement_date")):
            hire_date = doa or _derive_hire_date(dob, retirement_date)
        if hire_date is None and existing is not None:
            hire_date = existing.hire_date
        if hire_date is None:
            raise HTTPException(status_code=400, detail=f"hire_date missing in {source_label} and could not be derived.")

        if existing:
            retirement_target = _prefer_target(existing.retirement_date, retirement_date, has_col("retirement_date"))
            promo_role_target = _prefer_target(existing.promotion_role, promo_role, has_col("promotion_role"))
            promo_ready_target = _prefer_target(existing.promotion_ready_date, promo_ready, has_col("promotion_ready_date"))
            category_target = _prefer_target(existing.category, category, has_col("category"))
            pf_no_target = _prefer_target(existing.pf_no, pf_no, has_col("pf_no"))
            hrms_target = _prefer_target(existing.hrms, hrms, has_col("hrms"))
            crew_id_target = _prefer_target(existing.crew_id, crew_id, has_col("crew_id"))
            incoming_cli_id = str(get("cli_id")).strip() if has_col("cli_id") and get("cli_id") else None
            raw_cli_id = _prefer_target(existing.cli_id, incoming_cli_id, has_col("cli_id"))
            dob_target = _prefer_target(existing.dob, dob, has_col("dob"))
            doa_target = _prefer_target(existing.doa, doa, has_col("doa"))
            do_report_target = _prefer_target(existing.do_report, do_report, has_col("do_report"))
            status_target = _prefer_target(existing.status, status_val, has_col("status"))
            working_at_target = _prefer_target(existing.working_at, working_at, has_col("working_at") or working_at_override is not None)
            incoming_gradation = str(get("gradation")).strip() if has_col("gradation") and get("gradation") else None
            new_gradation = _prefer_target(existing.gradation, incoming_gradation, has_col("gradation"))
            incoming_cli = str(get("cli")).strip() if has_col("cli") and get("cli") else None
            raw_cli = _prefer_target(existing.cli, incoming_cli, has_col("cli"))
            existing_cli_clean, existing_cli_id_clean = _canonicalize_cli_name(
                existing.cli,
                existing.cli_id,
                canonical_by_id=canonical_by_id,
                alias_map=alias_map,
                id_by_name=id_by_name,
            )
            new_cli, cli_id_target = _canonicalize_cli_name(
                raw_cli,
                raw_cli_id,
                canonical_by_id=canonical_by_id,
                alias_map=alias_map,
                id_by_name=id_by_name,
            )
            if _cli_names_equivalent(existing_cli_clean, new_cli) and existing_cli_clean:
                new_cli = existing_cli_clean
            pme_due_target = _prefer_target(existing.pme_due, pme_due, has_col("pme_due"))
            technical_due_target = _prefer_target(existing.technical_due, technical_due, has_col("technical_due"))
            transportation_due_target = _prefer_target(existing.transportation_due, transportation_due, has_col("transportation_due"))
            field_updates = [
                ("Name", existing.name, _prefer_target(existing.name, str(name).strip(), has_col("name"))),
                ("Designation", existing.role, _prefer_target(existing.role, role, has_col("role"))),
                ("APPOINT DATE", existing.hire_date, _prefer_target(existing.hire_date, hire_date, has_col("hire_date") or has_col("doa") or has_col("dob") or has_col("retirement_date"))),
                ("Retirement Date", existing.retirement_date, retirement_target),
                ("Promotion Designation", existing.promotion_role, promo_role_target),
                ("Promotion Ready Date", existing.promotion_ready_date, promo_ready_target),
                ("Category", existing.category, category_target),
                ("PF No", existing.pf_no, pf_no_target),
                ("HRMS ID", existing.hrms, hrms_target),
                ("CREW ID", existing.crew_id, crew_id_target),
                ("CLI ID", existing_cli_id_clean, cli_id_target),
                ("DOB", existing.dob, dob_target),
                ("DOA", existing.doa, doa_target),
                ("DO Report", existing.do_report, do_report_target),
                ("Status", existing.status, status_target),
                ("Working At", existing.working_at, working_at_target),
                ("Gradation", existing.gradation, new_gradation),
                ("CLI", existing_cli_clean, new_cli),
                ("PME Due", existing.pme_due, pme_due_target),
                ("Technical Due", existing.technical_due, technical_due_target),
                ("Transportation Due", existing.transportation_due, transportation_due_target),
            ]
            changed_field_entries = [
                (label, old_value, new_value)
                for label, old_value, new_value in field_updates
                if old_value != new_value
            ]
            changed_fields = [
                f"{label}: {_format_sync_value(old_value)} -> {_format_sync_value(new_value)}"
                for label, old_value, new_value in changed_field_entries
            ]
            change_kind = _google_sync_change_label([label for label, _, _ in changed_field_entries])

            existing.name = _prefer_target(existing.name, str(name).strip(), has_col("name"))
            existing.role = _prefer_target(existing.role, role, has_col("role"))
            existing.hire_date = _prefer_target(existing.hire_date, hire_date, has_col("hire_date") or has_col("doa") or has_col("dob") or has_col("retirement_date"))
            existing.retirement_date = retirement_target
            existing.promotion_role = promo_role_target
            existing.promotion_ready_date = promo_ready_target
            existing.category = category_target
            existing.pf_no = pf_no_target
            existing.hrms = hrms_target
            existing.crew_id = crew_id_target
            existing.cli_id = cli_id_target
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
                    sync_details.append(f"{change_kind} {row_hint}: {'; '.join(changed_fields)}")
            elif sync_stats is not None:
                sync_stats["unchanged"] = sync_stats.get("unchanged", 0) + 1
        else:
            new_cli, new_cli_id = _canonicalize_cli_name(
                str(get("cli")).strip() if "cli" in col_index and get("cli") else None,
                str(get("cli_id")).strip() if "cli_id" in col_index and get("cli_id") else None,
                canonical_by_id=canonical_by_id,
                alias_map=alias_map,
                id_by_name=id_by_name,
            )
            employee = Employee(
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
                cli_id=new_cli_id,
                dob=dob,
                doa=doa,
                do_report=do_report,
                status=status_val,
                working_at=working_at,
                gradation=str(get("gradation")).strip() if "gradation" in col_index and get("gradation") else None,
                cli=new_cli,
                pme_due=pme_due,
                technical_due=technical_due,
                transportation_due=transportation_due,
            )
            session.add(employee)
            employees.append(employee)
            added += 1
            if sync_details is not None:
                sync_details.append(f"Added {row_hint}: Designation {_format_sync_value(role)}; Working At {_format_sync_value(working_at)}")

    if commit_changes:
        session.commit()
        _normalize_employee_cli_names(session)
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
        raise HTTPException(status_code=400, detail=f"Google Sheet tab '{requested_name}' was not found. Available tabs: {available}")
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


def _employees_match_smart_merge(left: Employee, right: Employee) -> bool:
    if left is right:
        return False
    left_name = _normalize_import_name(left.name)
    right_name = _normalize_import_name(right.name)
    if not left_name or not right_name or not _names_almost_same(left.name, right.name):
        return False
    left_role = normalize_role(left.role) if left.role else None
    right_role = normalize_role(right.role) if right.role else None
    if not left_role or left_role != right_role:
        return False
    if not left.dob or not right.dob or left.dob != right.dob:
        return False
    if left.retirement_date != right.retirement_date:
        return False

    left_cli_name, left_cli_id = _canonicalize_cli_name(left.cli, left.cli_id)
    right_cli_name, right_cli_id = _canonicalize_cli_name(right.cli, right.cli_id)
    if (left_cli_id or right_cli_id) and left_cli_id != right_cli_id:
        return False
    if (left_cli_name or right_cli_name) and not _cli_names_equivalent(left_cli_name, right_cli_name):
        return False

    left_pf = _emp_no_last5(left.pf_no)
    right_pf = _emp_no_last5(right.pf_no)
    if left_pf and right_pf and left_pf != right_pf:
        return False

    left_crew = _clean_import_text(left.crew_id)
    right_crew = _clean_import_text(right.crew_id)
    if left_crew and right_crew and left_crew != right_crew:
        return False

    left_hrms = _clean_import_text(left.hrms)
    right_hrms = _clean_import_text(right.hrms)
    if left_hrms and right_hrms and left_hrms != right_hrms:
        return False

    has_complementary_identifier = (
        (bool(left_crew) != bool(right_crew))
        or (bool(left_hrms) != bool(right_hrms))
    )
    return has_complementary_identifier


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
        "cli": _employee_cli_label(employee),
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
    for _, record in sorted(records.items()):
        payload.append(
            {
                "row_hint": str(record.get("row_hint") or ""),
                "name": _clean_import_text(record.get("name")) or "",
                "role": _clean_import_text(record.get("role")) or "",
                "pf_no": _clean_import_text(record.get("pf_no")) or "",
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


def _save_employee_master_service_snapshot(
    records: dict[str, dict[str, object]],
    crew_to_emp: dict[str, str],
) -> None:
    payload: dict[str, object] = {
        "records": [],
        "crew_to_emp": dict(sorted(crew_to_emp.items())),
    }
    serialized_records: list[dict[str, object]] = []
    for record_key, record in sorted(records.items()):
        serialized = _serialize_employee_payload(record)
        serialized["record_key"] = record_key
        serialized["row_hint"] = str(record.get("row_hint") or "")
        serialized["present_fields"] = sorted(str(field) for field in set(record.get("present_fields") or set()))
        serialized_records.append(serialized)
    payload["records"] = serialized_records
    EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_employee_master_service_snapshot() -> tuple[dict[str, dict[str, object]], dict[str, str]]:
    if not EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE.exists():
        return {}, {}
    try:
        raw = json.loads(EMPLOYEE_MASTER_SERVICE_SNAPSHOT_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}, {}
    if not isinstance(raw, dict):
        return {}, {}

    records_raw = raw.get("records")
    crew_to_emp_raw = raw.get("crew_to_emp")
    if not isinstance(records_raw, list) or not isinstance(crew_to_emp_raw, dict):
        return {}, {}

    records: dict[str, dict[str, object]] = {}
    for item in records_raw:
        if not isinstance(item, dict):
            continue
        record = _deserialize_employee_payload(item)
        emp_no = _clean_import_text(record.get("pf_no"))
        crew_id = _clean_import_text(record.get("crew_id"))
        record_key = str(item.get("record_key") or "").strip() or emp_no or (f"crew:{crew_id}" if crew_id else "")
        if not record_key:
            continue
        row_hint = str(item.get("row_hint") or record.get("name") or emp_no)
        present_fields_raw = item.get("present_fields")
        present_fields = {
            str(field)
            for field in present_fields_raw
            if isinstance(present_fields_raw, list) and field not in (None, "")
        }
        record["row_hint"] = row_hint
        record["present_fields"] = present_fields or {
            "name",
            "role",
            "pf_no",
            "crew_id",
            "dob",
            "hire_date",
            "doa",
            "retirement_date",
            "promotion_ready_date",
        }
        records[record_key] = record

    crew_to_emp = {
        _clean_import_text(crew_id) or "": _clean_import_text(emp_no) or ""
        for crew_id, emp_no in crew_to_emp_raw.items()
        if _clean_import_text(crew_id) and _clean_import_text(emp_no)
    }
    return records, crew_to_emp


def _save_employee_master_mismatch_actions(actions: list[dict[str, object]]) -> None:
    EMPLOYEE_MASTER_MISMATCH_ACTIONS_FILE.write_text(
        json.dumps(actions, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_employee_master_mismatch_actions() -> list[dict[str, object]]:
    if not EMPLOYEE_MASTER_MISMATCH_ACTIONS_FILE.exists():
        return []
    try:
        raw = json.loads(EMPLOYEE_MASTER_MISMATCH_ACTIONS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []
    if not isinstance(raw, list):
        return []
    actions: list[dict[str, object]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        normalized = dict(item)
        action_key = str(normalized.get("action_key") or "").strip()
        if not action_key:
            existing_id = normalized.get("existing_id")
            incoming_json = str(normalized.get("incoming_json") or "")
            normalized["action_key"] = f"{existing_id}:{hashlib.sha256(incoming_json.encode('utf-8')).hexdigest()}"
        actions.append(normalized)
    return actions


def _save_cli_nomination_mismatch_actions(actions: list[dict[str, object]]) -> None:
    CLI_NOMINATION_MISMATCH_ACTIONS_FILE.write_text(
        json.dumps(actions, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_cli_nomination_mismatch_actions() -> list[dict[str, object]]:
    if not CLI_NOMINATION_MISMATCH_ACTIONS_FILE.exists():
        return []
    try:
        raw = json.loads(CLI_NOMINATION_MISMATCH_ACTIONS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []
    if not isinstance(raw, list):
        return []
    actions: list[dict[str, object]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        normalized = dict(item)
        action_key = str(normalized.get("action_key") or "").strip()
        if not action_key:
            existing_id = normalized.get("existing_id")
            incoming_json = str(normalized.get("incoming_json") or "")
            normalized["action_key"] = f"{existing_id}:{hashlib.sha256(incoming_json.encode('utf-8')).hexdigest()}"
        actions.append(normalized)
    return actions


def _remove_cli_nomination_mismatch_action(action_key: str) -> list[dict[str, object]]:
    remaining = [
        item
        for item in _load_cli_nomination_mismatch_actions()
        if str(item.get("action_key") or "") != action_key
    ]
    _save_cli_nomination_mismatch_actions(remaining)
    return remaining


def _clear_cli_nomination_mismatch_actions() -> list[dict[str, object]]:
    _save_cli_nomination_mismatch_actions([])
    return []


def _remove_employee_master_mismatch_action(action_key: str) -> list[dict[str, object]]:
    remaining = [
        item
        for item in _load_employee_master_mismatch_actions()
        if str(item.get("action_key") or "") != action_key
    ]
    _save_employee_master_mismatch_actions(remaining)
    return remaining


def _split_employee_update_details(details: list[str]) -> tuple[list[str], list[str], list[str]]:
    added_details: list[str] = []
    updated_details: list[str] = []
    deduplicated_details: list[str] = []
    for item in details:
        if item.startswith("Added "):
            added_details.append(item)
        elif item.startswith("Updated "):
            updated_details.append(item)
        elif item.startswith("Deduplicated "):
            deduplicated_details.append(item)
    return added_details, updated_details, deduplicated_details


def _save_employee_master_review_report(
    *,
    update_notice: str = "",
    update_warning: str = "",
    update_details: Optional[list[str]] = None,
    warning_details: Optional[list[str]] = None,
    update_mismatch_actions: Optional[list[dict[str, object]]] = None,
    update_added_details: Optional[list[str]] = None,
    update_updated_details: Optional[list[str]] = None,
    update_deduplicated_details: Optional[list[str]] = None,
    append_history: bool = False,
) -> None:
    entry = {
        "saved_at": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "update_notice": update_notice,
        "update_warning": update_warning,
        "update_details": list(update_details or []),
        "warning_details": list(warning_details or []),
        "update_mismatch_actions": list(update_mismatch_actions or []),
        "update_added_details": list(update_added_details or []),
        "update_updated_details": list(update_updated_details or []),
        "update_deduplicated_details": list(update_deduplicated_details or []),
    }
    existing = _load_employee_master_review_report()
    history = list(existing.get("history") or [])
    if append_history:
        history = [entry] + history
    elif history:
        history[0] = entry
    else:
        history = [entry]
    payload = dict(entry)
    payload["history"] = history[:25]
    EMPLOYEE_MASTER_REVIEW_REPORT_FILE.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_employee_master_review_report() -> dict[str, object]:
    if not EMPLOYEE_MASTER_REVIEW_REPORT_FILE.exists():
        return {}
    try:
        raw = json.loads(EMPLOYEE_MASTER_REVIEW_REPORT_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(raw, dict):
        return {}
    history = raw.get("history")
    if not isinstance(history, list):
        entry = dict(raw)
        raw["history"] = [entry] if entry else []
    return raw


def _delete_employee_master_review_report_at(index: int) -> dict[str, object]:
    saved = _load_employee_master_review_report()
    history = list(saved.get("history") or [])
    if index < 0 or index >= len(history):
        raise HTTPException(status_code=404, detail="Review report not found.")
    del history[index]
    if not history:
        if EMPLOYEE_MASTER_REVIEW_REPORT_FILE.exists():
            EMPLOYEE_MASTER_REVIEW_REPORT_FILE.unlink()
        return {}
    latest = dict(history[0])
    latest["history"] = history[:25]
    EMPLOYEE_MASTER_REVIEW_REPORT_FILE.write_text(
        json.dumps(latest, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )
    return latest


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
        keep_payload = _cleanup_row_payload(keeper)
        remove_payloads = [_cleanup_row_payload(employee) for employee in remove_rows]
        plan.append(
            {
                "reason": reason,
                "keep": keep_payload,
                "remove": remove_payloads,
                "rows": [keep_payload] + remove_payloads,
                "row_ids": [employee.id for employee in ordered if employee.id is not None],
                "suggested_keep_id": keeper.id if keeper.id is not None else None,
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

    smart_merge_groups: dict[tuple[str, str, str, str, str], list[Employee]] = {}
    for employee in employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        name_key = _normalize_import_name(employee.name)
        role_key = normalize_role(employee.role) if employee.role else None
        dob_key = employee.dob.isoformat() if employee.dob else None
        retirement_key = employee.retirement_date.isoformat() if employee.retirement_date else ""
        cli_name, cli_id = _canonicalize_cli_name(employee.cli, employee.cli_id)
        cli_key = _cli_name_key(cli_name).upper() or (cli_id or "").upper()
        if not name_key or not role_key or not dob_key:
            continue
        smart_merge_groups.setdefault((name_key, role_key, dob_key, retirement_key, cli_key), []).append(employee)

    for rows in smart_merge_groups.values():
        if len(rows) < 2:
            continue
        matching_rows = [rows[0]]
        for employee in rows[1:]:
            if all(_employees_match_smart_merge(employee, existing) for existing in matching_rows):
                matching_rows.append(employee)
        if len(matching_rows) > 1:
            register_plan("Smart merge: same Name + Designation + DOB + Retirement + CLI with complementary IDs", matching_rows)

    remaining_employees = [
        employee
        for employee in employees
        if employee.id is None or employee.id not in used_ids
    ]
    for employee in remaining_employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        smart_group = [employee]
        for candidate in remaining_employees:
            if candidate is employee:
                continue
            if candidate.id is not None and candidate.id in used_ids:
                continue
            if all(_employees_match_smart_merge(candidate, existing) for existing in smart_group):
                smart_group.append(candidate)
        unique_group = []
        seen_group_ids: set[int] = set()
        for item in smart_group:
            item_id = item.id or 0
            if item_id in seen_group_ids:
                continue
            seen_group_ids.add(item_id)
            unique_group.append(item)
        if len(unique_group) > 1:
            register_plan("Smart merge: pairwise same Name + Designation + DOB + Retirement + CLI", unique_group)

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
        "name",
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
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
        "cli_id",
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
        "name",
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
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
        "cli_id",
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

    plan: list[dict[str, object]] = []
    conflicts: list[dict[str, object]] = []
    covered_review_ids: set[int] = set()

    for item in base_plan:
        item = {
            **item,
            "merge_action": "/uploads/employee-master-cleanup-merge",
            "delete_action": "/uploads/employee-master-cleanup-delete-row",
            "keep_action": "/uploads/employee-master-cleanup-keep-both",
            "keep_button_label": "Keep Both",
            "keep_id": item.get("suggested_keep_id"),
            "allow_merge": True,
            "allow_delete": True,
        }
        plan.append(item)
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
                "delete_action": "/uploads/employee-master-cleanup-delete-row",
                "keep_action": "/uploads/employee-master-cleanup-keep-both",
                "keep_button_label": "Keep Both",
                "keep_id": item.get("suggested_keep_id"),
                "allow_merge": True,
                "allow_delete": True,
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
    duplicate_ids: set[int] = set()

    def register_duplicate(candidate: Employee, reason: str) -> None:
        candidate_id = candidate.id or 0
        if candidate_id in duplicate_ids:
            return
        duplicate_ids.add(candidate_id)
        duplicates.append((candidate, reason))

    for employee in list(employees):
        if employee is target:
            continue

        candidate_pf = _clean_import_text(employee.pf_no)
        candidate_name = _normalize_import_name(employee.name)
        same_working_at = _working_at_key(employee.working_at) == target_working_at
        blank_vs_value_working_at = _one_working_at_blank(employee.working_at, target.working_at)
        candidate_last5 = _emp_no_last5(candidate_pf)

        if dob and target_last5 and employee.dob == dob and candidate_last5 == target_last5:
            register_duplicate(employee, "Same DOB + EMP NO last 5")
            continue

        if target_name and dob and candidate_name and _names_almost_same(employee.name, name) and employee.dob == dob:
            if same_working_at and (
                candidate_pf is None or emp_no is None or (target_last5 and candidate_last5 == target_last5)
            ):
                register_duplicate(employee, "Same Name + DOB")
                continue
            if blank_vs_value_working_at and target_last5 and candidate_pf and candidate_last5 == target_last5:
                register_duplicate(employee, "Same Name + DOB and one Working At is blank")
                continue

        if not same_working_at:
            continue

        if target_name and target_role and candidate_name and _names_almost_same(employee.name, name) and normalize_role(employee.role) == target_role:
            if candidate_pf and target_last5 and candidate_last5 == target_last5:
                register_duplicate(employee, "Same Name + Designation")
                continue

        if _employees_match_smart_merge(target, employee):
            register_duplicate(employee, "Smart merge: same Name + Designation + DOB + Retirement + CLI with complementary IDs")

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

    header_row = next(
        (
            row for row in rows
            if {
                "crewname",
                "crewdesg",
                "crewid",
                "empno",
                "birthdate",
                "appointdate",
                "retirementdate",
            }.issubset({_employee_norm(cell) for cell in row if cell not in (None, "", " ")})
        ),
        None,
    )
    if header_row is None:
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

        if not emp_no and not crew_id:
            warnings.append(f"Service Particulars {row_hint}: skipped because EMP NO and CREW ID are both blank.")
            continue
        if emp_no and any(_clean_import_text(item.get("pf_no")) == emp_no for item in records.values()):
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

        record_key = emp_no or f"crew:{crew_id}"
        records[record_key] = {
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
            crew_to_emp[crew_id] = record_key

    for emp_no in sorted(duplicate_emp):
        warnings.append(f"Service Particulars duplicate EMP NO skipped: {emp_no}")
        duplicate_keys = [key for key, value in records.items() if _clean_import_text(value.get("pf_no")) == emp_no]
        for key in duplicate_keys:
            records.pop(key, None)
    for crew_id in sorted(duplicate_crew):
        record_key = crew_to_emp.get(crew_id)
        if record_key:
            records.pop(record_key, None)
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
    row_dicts: list[dict[str, object]] = []
    normalized_header: dict[str, str] = {}

    try:
        wb = load_workbook(filename=BytesIO(content), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header_row = next(
            (row for row in rows if "crewid" in {_employee_norm(cell) for cell in row if cell not in (None, "")}),
            None,
        )
        if header_row is not None:
            headers = [_clean_import_text(cell) or "" for cell in header_row]
            normalized_header = {_employee_norm(name): name for name in headers if name}
            for row in rows[rows.index(header_row) + 1 :]:
                if not any(cell not in (None, "", " ") for cell in row):
                    continue
                row_dicts.append(
                    {
                        headers[idx]: row[idx] if idx < len(row) else None
                        for idx in range(len(headers))
                        if headers[idx]
                    }
                )
    except Exception:
        row_dicts = []
        normalized_header = {}

    if not row_dicts:
        text = _decode_uploaded_text(content)
        reader = csv.DictReader(StringIO(text))
        if not reader.fieldnames:
            raise HTTPException(status_code=400, detail="CMS other bio data file has no header row.")
        normalized_header = {_employee_norm(name): name for name in reader.fieldnames if name}
        row_dicts = list(reader)

    if "crewid" not in normalized_header:
        raise HTTPException(status_code=400, detail="CMS other bio data is missing column: CREWID")

    seen_hrms: set[str] = set()
    for row in row_dicts:
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


def _serialize_employee_payload(record_values: dict[str, object]) -> dict[str, object]:
    def _date_to_iso(value: object) -> str | None:
        return value.isoformat() if isinstance(value, date) else None

    return {
        "name": record_values.get("name"),
        "role": record_values.get("role"),
        "hire_date": _date_to_iso(record_values.get("hire_date")),
        "doa": _date_to_iso(record_values.get("doa")),
        "retirement_date": _date_to_iso(record_values.get("retirement_date")),
        "promotion_ready_date": _date_to_iso(record_values.get("promotion_ready_date")),
        "category": record_values.get("category"),
        "pf_no": record_values.get("pf_no"),
        "crew_id": record_values.get("crew_id"),
        "dob": _date_to_iso(record_values.get("dob")),
        "pme_due": _date_to_iso(record_values.get("pme_due")),
    }


def _deserialize_employee_payload(payload: dict[str, object]) -> dict[str, object]:
    def _iso_to_date(value: object) -> date | None:
        if isinstance(value, str) and value:
            try:
                return date.fromisoformat(value)
            except ValueError:
                return None
        return None

    return {
        "name": _clean_import_text(payload.get("name")),
        "role": normalize_role(_clean_import_text(payload.get("role")) or ""),
        "hire_date": _iso_to_date(payload.get("hire_date")),
        "doa": _iso_to_date(payload.get("doa")),
        "retirement_date": _iso_to_date(payload.get("retirement_date")),
        "promotion_ready_date": _iso_to_date(payload.get("promotion_ready_date")),
        "category": _clean_import_text(payload.get("category"), blank_na=True),
        "pf_no": _clean_import_text(payload.get("pf_no")),
        "crew_id": _clean_import_text(payload.get("crew_id")),
        "dob": _iso_to_date(payload.get("dob")),
        "pme_due": _iso_to_date(payload.get("pme_due")),
    }


def _build_mismatch_action(
    *,
    reason: str,
    row_hint: str,
    existing: Employee,
    incoming_values: dict[str, object],
) -> dict[str, object]:
    incoming_payload = _serialize_employee_payload(incoming_values)
    incoming_json = json.dumps(incoming_payload, default=str)
    return {
        "reason": reason,
        "row_hint": row_hint,
        "existing_id": existing.id,
        "action_key": f"{existing.id}:{hashlib.sha256(incoming_json.encode('utf-8')).hexdigest()}",
        "existing": {
            "name": existing.name,
            "role": existing.role,
            "pf_no": existing.pf_no,
            "crew_id": existing.crew_id,
            "dob": existing.dob.isoformat() if existing.dob else "",
            "working_at": existing.working_at or "",
        },
        "incoming": incoming_payload,
        "incoming_json": incoming_json,
    }


def _upsert_employee_master_records(
    session: Session,
    records: dict[str, dict[str, object]],
    warnings: list[str],
    sync_details: list[str],
) -> tuple[int, int, int, int, int, list[dict[str, object]]]:
    added = 0
    updated = 0
    unchanged = 0
    skipped = 0
    deduplicated = 0
    mismatch_actions: list[dict[str, object]] = []
    employees = session.exec(select(Employee)).all()

    by_pf: dict[str, list[Employee]] = {}
    by_pf_match: dict[str, list[Employee]] = {}
    by_crew_id: dict[str, list[Employee]] = {}

    def rebuild_exact_indexes() -> None:
        by_pf.clear()
        by_pf_match.clear()
        by_crew_id.clear()
        for employee in employees:
            pf_value = _clean_import_text(employee.pf_no)
            if pf_value:
                by_pf.setdefault(pf_value, []).append(employee)
                pf_match_value = _emp_no_match_key(pf_value)
                if pf_match_value:
                    by_pf_match.setdefault(pf_match_value, []).append(employee)
            crew_value = _clean_import_text(employee.crew_id)
            if crew_value:
                by_crew_id.setdefault(crew_value, []).append(employee)

    rebuild_exact_indexes()

    for record_key, record in records.items():
        emp_no = _clean_import_text(record.get("pf_no"))
        row_hint = str(record.get("row_hint") or record.get("name") or emp_no or record_key)
        name = _clean_import_text(record.get("name"))
        role = _clean_import_text(record.get("role"))
        hire_date = record.get("hire_date")
        present_fields = set(record.get("present_fields") or set())

        if not name or not role or not isinstance(hire_date, date):
            warnings.append(f"{row_hint}: skipped because name, designation, or appoint date is missing after merge.")
            skipped += 1
            continue

        pf_matches: list[Employee] = []
        if emp_no:
            pf_matches = by_pf.get(emp_no, [])
            if not pf_matches:
                pf_matches = by_pf_match.get(_emp_no_match_key(emp_no) or "", [])
        if emp_no and len(pf_matches) > 1:
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
                if emp_no and existing.pf_no and not _emp_no_matches(existing.pf_no, emp_no):
                    warnings.append(
                        f"{row_hint}: skipped because EMP NO {emp_no} conflicts with existing employee EMP NO {existing.pf_no}."
                    )
                    mismatch_actions.append(
                        _build_mismatch_action(
                            reason="EMP NO conflicts with existing employee",
                            row_hint=row_hint,
                            existing=existing,
                            incoming_values={
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
                            },
                        )
                    )
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
    return added, updated, unchanged, skipped, deduplicated, mismatch_actions


@app.post("/uploads/employee-master-sync")
async def upload_employee_master_sync(
    request: Request,
    service_file: Optional[UploadFile] = File(None),
    cms_file: UploadFile = File(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        service_name = service_file.filename if service_file is not None else ""
        cms_name = cms_file.filename or ""
        if not cms_name.lower().endswith((".csv", ".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="CMS other bio data file must be a .csv or .xlsx workbook.")

        cms_content = await cms_file.read()
        warnings: list[str] = []
        sync_details: list[str] = []
        service_uploaded = bool(service_file is not None and service_name.strip())
        if service_uploaded:
            if not service_name.lower().endswith((".xlsx", ".xlsm")):
                raise HTTPException(status_code=400, detail="Service Particulars file must be an .xlsx workbook.")
            service_content = await service_file.read()
            records, hrms_to_emp = _build_service_particular_records(service_content, warnings)
            _save_employee_master_service_snapshot(records, hrms_to_emp)
            sync_details.append("Loaded fresh Service Particulars reference from uploaded file.")
        else:
            records, hrms_to_emp = _load_employee_master_service_snapshot()
            if not records:
                raise HTTPException(
                    status_code=400,
                    detail="No saved Service Particulars reference found. Upload a Service Particulars file once first.",
                )
            sync_details.append("Loaded saved Service Particulars reference from previous upload.")
        _merge_cms_other_bio(records, hrms_to_emp, cms_content, warnings)
        _save_employee_master_source_snapshot(records)
        added, updated, unchanged, skipped, deduplicated, mismatch_actions = _upsert_employee_master_records(
            session,
            records,
            warnings,
            sync_details,
        )
        _save_employee_master_mismatch_actions(mismatch_actions)

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
        added_details, updated_details, deduplicated_details = _split_employee_update_details(sync_details)
        _save_employee_master_review_report(
            update_notice=notice,
            update_warning=warning_text,
            update_details=sync_details,
            warning_details=warnings,
            update_mismatch_actions=mismatch_actions,
            update_added_details=added_details,
            update_updated_details=updated_details,
            update_deduplicated_details=deduplicated_details,
            append_history=True,
        )

        return _uploads_template_response(
            request,
            update_notice=notice,
            update_warning=warning_text,
            update_details=sync_details,
            warning_details=warnings,
            update_mismatch_actions=mismatch_actions,
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Employee table update failed."
        return _uploads_template_response(request, update_error=detail, status_code=exc.status_code)
    except Exception as exc:
        return _uploads_template_response(request, update_error=str(exc), status_code=500)


def _create_employee_from_payload(payload: dict[str, object]) -> Employee:
    values = _deserialize_employee_payload(payload)
    name = _clean_import_text(values.get("name"))
    role = _clean_import_text(values.get("role"))
    hire_date = values.get("hire_date")
    if not name or not role or not isinstance(hire_date, date):
        raise ValueError("Incoming row is missing name, designation, or appoint date.")
    return Employee(
        name=name,
        role=normalize_role(role),
        hire_date=hire_date,
        retirement_date=values.get("retirement_date"),
        promotion_ready_date=values.get("promotion_ready_date"),
        category=values.get("category"),
        pf_no=values.get("pf_no"),
        crew_id=values.get("crew_id"),
        dob=values.get("dob"),
        doa=values.get("doa"),
        pme_due=values.get("pme_due"),
        status="ACTIVE",
    )


@app.post("/uploads/employee-master-service-reference")
async def upload_employee_master_service_reference(
    request: Request,
    service_file: UploadFile = File(...),
    action_password: str = Form(...),
):
    try:
        _validate_sensitive_action_password(action_password)
        service_name = service_file.filename or ""
        if not service_name.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Service Particulars file must be an .xlsx workbook.")
        service_content = await service_file.read()
        warnings: list[str] = []
        records, crew_to_emp = _build_service_particular_records(service_content, warnings)
        _save_employee_master_service_snapshot(records, crew_to_emp)
        notice = f"Service Particulars reference saved: {len(records)} employee rows ready for later CMS uploads."
        warning_text = f"Issues found while saving reference: {len(warnings)}" if warnings else ""
        return _uploads_template_response(
            request,
            update_notice=notice,
            update_warning=warning_text,
            warning_details=warnings,
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Service Particulars reference save failed."
        return _uploads_template_response(request, update_error=detail, status_code=exc.status_code)
    except Exception as exc:
        return _uploads_template_response(request, update_error=str(exc), status_code=500)


@app.post("/uploads/employee-master-mismatch-merge")
async def upload_employee_master_mismatch_merge(
    request: Request,
    existing_id: int = Form(...),
    incoming_payload: str = Form(...),
    action_key: str = Form(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        incoming_data = json.loads(incoming_payload)
        existing = session.get(Employee, existing_id)
        if not existing:
            raise HTTPException(status_code=404, detail="Existing row not found.")
        employee = _create_employee_from_payload(incoming_data)
        session.add(employee)
        session.commit()
        mismatch_actions = _remove_employee_master_mismatch_action(action_key)
        notice = "Mismatch merge complete: incoming row added alongside existing."
        saved_report = _load_employee_master_review_report()
        _save_employee_master_review_report(
            update_notice=notice,
            update_warning=str(saved_report.get("update_warning") or ""),
            update_details=list(saved_report.get("update_details") or []),
            warning_details=list(saved_report.get("warning_details") or []),
            update_mismatch_actions=mismatch_actions,
            update_added_details=list(saved_report.get("update_added_details") or []),
            update_updated_details=list(saved_report.get("update_updated_details") or []),
            update_deduplicated_details=list(saved_report.get("update_deduplicated_details") or []),
        )
        return _uploads_template_response(request, update_notice=notice, update_mismatch_actions=mismatch_actions)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Merge failed."
        return _uploads_template_response(request, update_error=detail, status_code=exc.status_code)
    except Exception as exc:
        return _uploads_template_response(request, update_error=str(exc), status_code=500)


@app.post("/uploads/employee-master-mismatch-delete")
async def upload_employee_master_mismatch_delete(
    request: Request,
    existing_id: int = Form(...),
    incoming_payload: str = Form(...),
    action_key: str = Form(...),
    delete_target: str = Form(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        incoming_data = json.loads(incoming_payload)
        existing = session.get(Employee, existing_id)
        if not existing:
            raise HTTPException(status_code=404, detail="Existing row not found.")
        if delete_target == "existing":
            session.delete(existing)
            employee = _create_employee_from_payload(incoming_data)
            session.add(employee)
            session.commit()
            notice = "Delete complete: existing row removed, incoming row kept."
        else:
            notice = "Delete complete: incoming row ignored, existing row kept."
        mismatch_actions = _remove_employee_master_mismatch_action(action_key)
        saved_report = _load_employee_master_review_report()
        _save_employee_master_review_report(
            update_notice=notice,
            update_warning=str(saved_report.get("update_warning") or ""),
            update_details=list(saved_report.get("update_details") or []),
            warning_details=list(saved_report.get("warning_details") or []),
            update_mismatch_actions=mismatch_actions,
            update_added_details=list(saved_report.get("update_added_details") or []),
            update_updated_details=list(saved_report.get("update_updated_details") or []),
            update_deduplicated_details=list(saved_report.get("update_deduplicated_details") or []),
        )
        return _uploads_template_response(request, update_notice=notice, update_mismatch_actions=mismatch_actions)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Delete failed."
        return _uploads_template_response(request, update_error=detail, status_code=exc.status_code)
    except Exception as exc:
        return _uploads_template_response(request, update_error=str(exc), status_code=500)


@app.post("/uploads/employee-master-review-delete")
async def upload_employee_master_review_delete(
    request: Request,
    report_index: int = Form(...),
    action_password: str = Form(...),
):
    try:
        _validate_sensitive_action_password(action_password)
        updated = _delete_employee_master_review_report_at(report_index)
        history = list(updated.get("history") or [])
        notice = "Review report deleted."
        if not history:
            notice = "Review report deleted. No saved reports remaining."
        return templates.TemplateResponse("uploads.html", _uploads_context(request, update_notice=notice), status_code=200)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Review report delete failed."
        return _uploads_template_response(request, update_error=detail, status_code=exc.status_code)
    except Exception as exc:
        return _uploads_template_response(request, update_error=str(exc), status_code=500)


def _normalize_li_grading_header(value: object | None) -> str:
    text = _clean_import_text(value)
    if text is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "", text.upper())


def _load_cli_upload_rows(content: bytes) -> list[tuple[object, ...]]:
    workbook = load_workbook(filename=BytesIO(content), data_only=True)
    worksheet = workbook.active
    return list(worksheet.iter_rows(values_only=True))


def _detect_cli_upload_kind(rows: list[tuple[object, ...]]) -> str:
    for row in rows[:8]:
        normalized = {_normalize_li_grading_header(cell) for cell in row if _normalize_li_grading_header(cell)}
        if not normalized:
            continue
        if {"CLIID", "CLINAME", "ALLOTEDDESIG"}.issubset(normalized):
            return "cli_matrix"
        if {"CLIID", "NAME", "EMPNO", "DESIG"}.issubset(normalized):
            return "cli_biodata"
        if {"CLIID", "CLINAME", "CREWID", "NAME", "CURRENTGRADE", "DUEDATE"}.issubset(normalized):
            return "li_grading"
        if {"LIID", "CREWID", "CREWNAME", "GRADE", "GRADINGDATE"}.issubset(normalized):
            return "li_grading"
    return "unknown"


def _parse_li_grading_workbook(content: bytes) -> tuple[list[dict[str, object]], list[str]]:
    rows = _load_cli_upload_rows(content)
    if not rows:
        raise HTTPException(status_code=400, detail="CLI Grading workbook is empty.")

    header_row_index: int | None = None
    cli_id_idx: int | None = None
    cli_name_idx: int | None = None
    crew_idx: int | None = None
    name_idx: int | None = None
    role_idx: int | None = None
    current_grade_idx: int | None = None
    due_date_idx: int | None = None
    alternate_layout = False

    for idx, row in enumerate(rows):
        normalized = [_normalize_li_grading_header(cell) for cell in row]
        if "CLIID" not in normalized or "CLINAME" not in normalized or "CREWID" not in normalized or "NAME" not in normalized or "CURRENTGRADE" not in normalized:
            if {"LIID", "CREWID", "CREWNAME", "GRADE", "GRADINGDATE"}.issubset(normalized):
                alternate_layout = True
                header_row_index = idx
                cli_id_idx = normalized.index("LIID")
                cli_name_idx = None
                crew_idx = normalized.index("CREWID")
                name_idx = normalized.index("CREWNAME")
                role_idx = None
                current_grade_idx = normalized.index("GRADE")
                due_date_idx = normalized.index("GRADINGDATE")
                break
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
        cli_id_idx = normalized.index("CLIID")
        cli_name_idx = normalized.index("CLINAME")
        crew_idx = normalized.index("CREWID")
        name_idx = normalized.index("NAME")
        header_row_index = idx
        break

    if header_row_index is None or cli_id_idx is None or crew_idx is None or name_idx is None or current_grade_idx is None or due_date_idx is None:
        raise HTTPException(
            status_code=400,
            detail="Could not find the CLI Grading columns. Required columns: CLI ID, CLI NAME, CREW ID, NAME, DESIG., CURRENT GRADE, DUE DATE. Alternate supported format: LI ID, CREW ID, CREW NAME, GRADE, GRADING DATE.",
        )

    warnings: list[str] = []
    records: list[dict[str, object]] = []

    for row_number, row in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        def get(column_index: int | None) -> object | None:
            if column_index is None or column_index >= len(row):
                return None
            return row[column_index]

        cli_id = _clean_import_text(get(cli_id_idx))
        cli_name = _clean_import_text(get(cli_name_idx))
        crew_id = _clean_import_text(get(crew_idx))
        name = _clean_import_text(get(name_idx))
        role_raw = _clean_import_text(get(role_idx))
        current_grade = _clean_import_text(get(current_grade_idx))
        due_raw = get(due_date_idx)

        if not any([cli_id, cli_name, crew_id, name, role_raw, current_grade, due_raw]):
            continue

        row_hint = name or crew_id or f"row {row_number}"
        if not cli_id:
            warnings.append(f"CLI Grading {row_hint}: skipped because CLI ID is blank.")
            continue
        if not cli_name and not alternate_layout:
            warnings.append(f"CLI Grading {row_hint}: skipped because CLI NAME is blank.")
            continue
        if not name:
            warnings.append(f"CLI Grading row {row_number}: skipped because NAME is blank.")
            continue
        if not role_raw and not alternate_layout:
            warnings.append(f"CLI Grading {row_hint}: skipped because DESIG. is blank.")
            continue
        if not current_grade:
            warnings.append(f"CLI Grading {row_hint}: skipped because CURRENT GRADE is blank.")
            continue

        try:
            due_date = _excel_to_date_with_correction(due_raw, warnings, "CLI Grading", row_hint, "due_date") if due_raw not in (None, "") else None
        except ValueError as exc:
            warnings.append(f"CLI Grading {row_hint}: skipped because DUE DATE is invalid ({exc}).")
            continue

        records.append(
            {
                "cli_id": cli_id,
                "cli_name": cli_name,
                "crew_id": crew_id,
                "name": name,
                "role": normalize_role(role_raw) if role_raw else "",
                "gradation": current_grade.upper(),
                "grading_due": due_date,
                "row_hint": row_hint,
            }
        )

    if not records:
        raise HTTPException(status_code=400, detail="CLI Grading workbook did not produce any usable rows.")
    return records, warnings


def _parse_cli_biodata_workbook(content: bytes) -> tuple[list[dict[str, object]], list[str]]:
    rows = _load_cli_upload_rows(content)
    if not rows:
        raise HTTPException(status_code=400, detail="CLITI Biodata workbook is empty.")

    header_row_index: int | None = None
    cli_id_idx: int | None = None
    name_idx: int | None = None
    emp_no_idx: int | None = None
    role_idx: int | None = None
    mobile_idx: int | None = None
    dob_idx: int | None = None
    doa_idx: int | None = None
    dop_idx: int | None = None
    hq_idx: int | None = None

    for idx, row in enumerate(rows):
        normalized = [_normalize_li_grading_header(cell) for cell in row]
        if "CLIID" not in normalized or "NAME" not in normalized or "EMPNO" not in normalized:
            continue
        role_idx = next((i for i, value in enumerate(normalized) if value in {"DESIG", "DESIGNATION", "ROLE"}), None)
        if role_idx is None:
            continue
        header_row_index = idx
        cli_id_idx = normalized.index("CLIID")
        name_idx = normalized.index("NAME")
        emp_no_idx = normalized.index("EMPNO")
        mobile_idx = next(
            (
                i
                for i, value in enumerate(normalized)
                if value in {
                    "MOBILE",
                    "MOBILENO",
                    "MOBILENUMBER",
                    "MOBNO",
                    "MOBILENUM",
                    "MOB",
                    "PHONE",
                    "PHONENO",
                    "PHONENUMBER",
                    "CONTACTNO",
                    "CONTACTNUMBER",
                    "CELLNO",
                    "WHATSAPPNO",
                }
            ),
            None,
        )
        dob_idx = next((i for i, value in enumerate(normalized) if value in {"DOB", "DOBSTAR"}), None)
        doa_idx = next((i for i, value in enumerate(normalized) if value in {"DOA", "DOASTAR"}), None)
        dop_idx = next((i for i, value in enumerate(normalized) if value in {"DOP", "DOPSTAR"}), None)
        hq_idx = next((i for i, value in enumerate(normalized) if value in {"HQ", "WORKINGAT"}), None)
        break

    if header_row_index is None or None in {cli_id_idx, name_idx, emp_no_idx, role_idx}:
        raise HTTPException(
            status_code=400,
            detail="Could not find the CLITI Biodata columns. Required columns: CLI ID, NAME, EMP NO, DESIG.",
        )

    warnings: list[str] = []
    records: list[dict[str, object]] = []

    for row_number, row in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        def get(column_index: int | None) -> object | None:
            if column_index is None or column_index >= len(row):
                return None
            return row[column_index]

        cli_id = _clean_import_text(get(cli_id_idx))
        name = _clean_import_text(get(name_idx))
        emp_no = _clean_import_text(get(emp_no_idx))
        role_raw = _clean_import_text(get(role_idx))
        mobile_no = _clean_import_text(get(mobile_idx))
        hq = _clean_import_text(get(hq_idx))

        if not any([cli_id, name, emp_no, role_raw, hq]):
            continue

        row_hint = name or cli_id or emp_no or f"row {row_number}"
        if not cli_id:
            warnings.append(f"CLITI Biodata {row_hint}: skipped because CLI ID is blank.")
            continue
        if not name:
            warnings.append(f"CLITI Biodata row {row_number}: skipped because NAME is blank.")
            continue
        if not role_raw:
            warnings.append(f"CLITI Biodata {row_hint}: skipped because DESIG is blank.")
            continue
        if not emp_no:
            warnings.append(f"CLITI Biodata {row_hint}: skipped because EMP NO is blank.")
            continue

        try:
            dob = _excel_to_date_with_correction(get(dob_idx), warnings, "CLITI Biodata", row_hint, "dob") if get(dob_idx) not in (None, "") else None
            doa = _excel_to_date_with_correction(get(doa_idx), warnings, "CLITI Biodata", row_hint, "doa") if get(doa_idx) not in (None, "") else None
            dop = _excel_to_date_with_correction(get(dop_idx), warnings, "CLITI Biodata", row_hint, "dop") if get(dop_idx) not in (None, "") else None
        except ValueError as exc:
            warnings.append(f"CLITI Biodata {row_hint}: skipped because a date is invalid ({exc}).")
            continue

        hire_date = doa or dop
        if hire_date is None:
            warnings.append(f"CLITI Biodata {row_hint}: skipped because DOA/DOP is blank.")
            continue

        records.append(
            {
                "cli_id": cli_id,
                "cli_name": name,
                "emp_no": emp_no,
                "name": name,
                "role": normalize_role(role_raw) or role_raw,
                "mobile_no": mobile_no,
                "dob": dob,
                "doa": doa,
                "do_report": dop,
                "hire_date": hire_date,
                "working_at": hq,
                "row_hint": row_hint,
            }
        )

    if not records:
        raise HTTPException(status_code=400, detail="CLITI Biodata workbook did not produce any usable rows.")
    return records, warnings


def _parse_cli_nomination_workbook(content: bytes) -> tuple[list[dict[str, object]], list[str]]:
    rows = _load_cli_upload_rows(content)
    if not rows:
        raise HTTPException(status_code=400, detail="CLI nomination workbook is empty.")

    header_row_index: int | None = None
    crew_idx: int | None = None
    name_idx: int | None = None
    hrms_idx: int | None = None
    cli_name_idx: int | None = None

    for idx, row in enumerate(rows):
        normalized = [_normalize_li_grading_header(cell) for cell in row]
        if {"CREWID", "CREWNAME", "HRMSID", "CLINAME"}.issubset(set(normalized)):
            header_row_index = idx
            crew_idx = normalized.index("CREWID")
            name_idx = normalized.index("CREWNAME")
            hrms_idx = normalized.index("HRMSID")
            cli_name_idx = normalized.index("CLINAME")
            break

    if header_row_index is None or None in {crew_idx, name_idx, hrms_idx, cli_name_idx}:
        raise HTTPException(
            status_code=400,
            detail="Could not find the CLI nomination columns. Required columns: CREWID, CREW NAME, HRMS ID, CLI Name.",
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
        hrms = _clean_import_text(get(hrms_idx))
        cli_name = _clean_import_text(get(cli_name_idx))

        if not any([crew_id, name, hrms, cli_name]):
            continue

        row_hint = name or crew_id or hrms or f"row {row_number}"
        if not cli_name:
            warnings.append(f"CLI nomination {row_hint}: skipped because CLI Name is blank.")
            continue
        if not crew_id and not hrms and not name:
            warnings.append(f"CLI nomination {row_hint}: skipped because CREW ID, HRMS ID, and CREW NAME are blank.")
            continue

        records.append(
            {
                "crew_id": crew_id,
                "name": name,
                "hrms": hrms,
                "cli_name": cli_name,
                "row_hint": row_hint,
            }
        )

    if not records:
        raise HTTPException(status_code=400, detail="CLI nomination workbook did not produce any usable rows.")
    return records, warnings


def _build_cli_nomination_mismatch_action(
    *,
    reason: str,
    row_hint: str,
    existing: Employee,
    incoming_values: dict[str, object],
) -> dict[str, object]:
    incoming_json = json.dumps(incoming_values, default=str)
    return {
        "reason": reason,
        "row_hint": row_hint,
        "existing_id": existing.id,
        "action_key": f"{existing.id}:{hashlib.sha256(incoming_json.encode('utf-8')).hexdigest()}",
        "existing": {
            "name": existing.name,
            "role": existing.role,
            "crew_id": existing.crew_id,
            "hrms": existing.hrms,
            "cli": existing.cli,
        },
        "incoming": incoming_values,
        "incoming_json": incoming_json,
    }


def _apply_cli_nomination_records(
    session: Session,
    records: list[dict[str, object]],
    warnings: list[str],
    source_filename: str = "",
) -> tuple[str, str, list[str], list[str], list[dict[str, object]]]:
    employees = session.exec(select(Employee)).all()
    canonical_by_id, alias_map, id_by_name = _build_cli_name_maps((employee.cli, employee.cli_id) for employee in employees)

    by_crew: dict[str, list[Employee]] = {}
    by_hrms: dict[str, list[Employee]] = {}
    by_name: dict[str, list[Employee]] = {}
    for employee in employees:
        crew_key = _clean_import_text(employee.crew_id)
        hrms_key = _clean_import_text(employee.hrms)
        name_key = _normalize_import_name(employee.name)
        if crew_key:
            by_crew.setdefault(crew_key.upper(), []).append(employee)
        if hrms_key:
            by_hrms.setdefault(hrms_key.upper(), []).append(employee)
        if name_key:
            by_name.setdefault(name_key, []).append(employee)

    updated = 0
    unchanged = 0
    skipped = 0
    details: list[str] = []
    mismatch_actions: list[dict[str, object]] = []
    touched_ids: set[int] = set()

    for record in records:
        row_hint = str(record.get("row_hint") or "")
        crew_key = (_clean_import_text(record.get("crew_id")) or "").upper()
        hrms_key = (_clean_import_text(record.get("hrms")) or "").upper()
        name_key = _normalize_import_name(record.get("name"))
        cli_name, cli_id = _canonicalize_cli_name(record.get("cli_name"), None, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)

        crew_matches = by_crew.get(crew_key, []) if crew_key else []
        hrms_matches = by_hrms.get(hrms_key, []) if hrms_key else []
        if crew_key and len(crew_matches) > 1:
            warnings.append(f"CLI nomination {row_hint}: skipped because CREW ID {crew_key} matched multiple employees.")
            skipped += 1
            continue
        if hrms_key and len(hrms_matches) > 1:
            warnings.append(f"CLI nomination {row_hint}: skipped because HRMS ID {hrms_key} matched multiple employees.")
            skipped += 1
            continue

        crew_target = crew_matches[0] if len(crew_matches) == 1 else None
        hrms_target = hrms_matches[0] if len(hrms_matches) == 1 else None
        target: Employee | None = None

        if (
            crew_target
            and hrms_target
            and crew_target.id != hrms_target.id
            and not _clean_import_text(crew_target.hrms)
            and name_key
            and _normalize_import_name(crew_target.name) == name_key
        ):
            hrms_target = None

        if crew_target and hrms_target and crew_target.id != hrms_target.id:
            mismatch_actions.append(
                _build_cli_nomination_mismatch_action(
                    reason="CREW ID and HRMS ID point to different employees",
                    row_hint=row_hint,
                    existing=crew_target,
                    incoming_values={
                        "name": _clean_import_text(record.get("name")),
                        "crew_id": _clean_import_text(record.get("crew_id")),
                        "hrms": _clean_import_text(record.get("hrms")),
                        "cli": cli_name,
                        "cli_id": cli_id,
                    },
                )
            )
            skipped += 1
            continue

        target = crew_target or hrms_target
        if target is None and name_key:
            name_matches = by_name.get(name_key, [])
            if len(name_matches) == 1:
                target = name_matches[0]
            elif len(name_matches) > 1:
                warnings.append(f"CLI nomination {row_hint}: skipped because CREW NAME matched multiple employees.")
                skipped += 1
                continue

        if target is None:
            warnings.append(f"CLI nomination {row_hint}: no matching employee row found.")
            skipped += 1
            continue

        if target.id is not None and target.id in touched_ids:
            warnings.append(f"CLI nomination {row_hint}: skipped because that employee already received a nomination update from another row in this workbook.")
            skipped += 1
            continue

        incoming_crew = _clean_import_text(record.get("crew_id"))
        incoming_hrms = _clean_import_text(record.get("hrms"))
        if incoming_crew and target.crew_id and _clean_import_text(target.crew_id) != incoming_crew and crew_target is None:
            mismatch_actions.append(
                _build_cli_nomination_mismatch_action(
                    reason="CREW ID conflicts with existing employee row",
                    row_hint=row_hint,
                    existing=target,
                    incoming_values={
                        "name": _clean_import_text(record.get("name")),
                        "crew_id": incoming_crew,
                        "hrms": incoming_hrms,
                        "cli": cli_name,
                        "cli_id": cli_id,
                    },
                )
            )
            skipped += 1
            continue
        if incoming_hrms and target.hrms and _clean_import_text(target.hrms) != incoming_hrms and hrms_target is None:
            mismatch_actions.append(
                _build_cli_nomination_mismatch_action(
                    reason="HRMS ID conflicts with existing employee row",
                    row_hint=row_hint,
                    existing=target,
                    incoming_values={
                        "name": _clean_import_text(record.get("name")),
                        "crew_id": incoming_crew,
                        "hrms": incoming_hrms,
                        "cli": cli_name,
                        "cli_id": cli_id,
                    },
                )
            )
            skipped += 1
            continue

        changes: list[str] = []
        if not target.crew_id and incoming_crew:
            changes.append(f"CREW ID: {_format_sync_value(target.crew_id)} -> {_format_sync_value(incoming_crew)}")
            target.crew_id = incoming_crew
        if not target.hrms and incoming_hrms:
            changes.append(f"HRMS ID: {_format_sync_value(target.hrms)} -> {_format_sync_value(incoming_hrms)}")
            target.hrms = incoming_hrms
        old_cli, old_cli_id = _canonicalize_cli_name(target.cli, target.cli_id, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)
        if old_cli != cli_name:
            changes.append(f"CLI: {_format_sync_value(old_cli)} -> {_format_sync_value(cli_name)}")
        if old_cli_id != cli_id:
            changes.append(f"CLI ID: {_format_sync_value(old_cli_id)} -> {_format_sync_value(cli_id)}")
        target.cli, target.cli_id = _canonicalize_cli_name(cli_name, cli_id, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)

        if changes:
            updated += 1
            details.append(f"Updated {target.name} ({target.crew_id or target.hrms or target.id}): " + "; ".join(changes))
        else:
            unchanged += 1
        if target.id is not None:
            touched_ids.add(target.id)

    session.commit()
    _normalize_employee_cli_names(session)
    _save_cli_nomination_mismatch_actions(mismatch_actions)
    notice_parts = []
    if updated:
        notice_parts.append(f"{updated} updated")
    if skipped:
        notice_parts.append(f"{skipped} skipped")
    if not notice_parts:
        notice_parts.append("No change found")
    notice = "CLI nomination import complete: " + ", ".join(notice_parts) + "."
    warning_message = f"Mismatch / auto-fixed records: {len(warnings)}" if warnings else ""
    return notice, warning_message, details, warnings, mismatch_actions


def _apply_cli_biodata_records(
    session: Session,
    records: list[dict[str, object]],
    warnings: list[str],
    source_filename: str = "",
) -> tuple[str, str, list[str], list[str]]:
    init_db()
    _refresh_cli_bio_reference_from_records(session, records, source_filename)
    details: list[str] = []
    session.commit()
    unique_count = len({str(record.get("cli_id") or "").strip() for record in records if str(record.get("cli_id") or "").strip()})
    if source_filename:
        details.append(f"CLI Bio Reference refreshed from {source_filename}.")
    details.append(f"Saved {unique_count} CLI bio reference rows.")
    notice = f"CLITI Biodata import complete: CLI Bio Reference refreshed with {unique_count} rows."
    warning_message = f"Mismatch / auto-fixed records: {len(warnings)}" if warnings else ""
    return notice, warning_message, details, warnings


@app.post("/upload-li-grading")
async def upload_li_grading(
    request: Request,
    file: Optional[UploadFile] = File(None),
    bio_data_file: Optional[UploadFile] = File(None),
    cli_nomination_file: Optional[UploadFile] = File(None),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        init_db()
        _validate_sensitive_action_password(action_password)
        filename = file.filename if file is not None else ""
        if (
            not (file and filename)
            and not (bio_data_file and bio_data_file.filename)
            and not (cli_nomination_file and cli_nomination_file.filename)
        ):
            raise HTTPException(status_code=400, detail="Upload a CLI Grading workbook, CLITI Biodata file, or a CLI nomination file.")

        content = b""
        upload_kind = "unknown"
        if file and filename:
            if not filename.lower().endswith((".xlsx", ".xlsm")):
                raise HTTPException(status_code=400, detail="Upload the CLI Grading .xlsx workbook.")
            content = await file.read()
            upload_kind = _detect_cli_upload_kind(_load_cli_upload_rows(content))
        nomination_notice = ""
        nomination_warning = ""
        nomination_details: list[str] = []
        nomination_warnings: list[str] = []
        nomination_actions: list[dict[str, object]] = []
        if cli_nomination_file and cli_nomination_file.filename:
            nomination_name = cli_nomination_file.filename or ""
            if not nomination_name.lower().endswith((".xlsx", ".xlsm")):
                raise HTTPException(status_code=400, detail="Upload the CLI nomination file as .xlsx.")
            nomination_content = await cli_nomination_file.read()
            nomination_records, nomination_warnings = _parse_cli_nomination_workbook(nomination_content)
            nomination_notice, nomination_warning, nomination_details, nomination_warnings, nomination_actions = _apply_cli_nomination_records(
                session,
                nomination_records,
                nomination_warnings,
                nomination_name,
            )
        if upload_kind == "unknown" and nomination_notice:
            return templates.TemplateResponse(
                "cli.html",
                _cli_page_context(
                    request,
                    session,
                    grading_update_notice=nomination_notice,
                    grading_update_warning=nomination_warning,
                    grading_update_details=nomination_details,
                    grading_warning_details=nomination_warnings,
                    nomination_mismatch_actions=nomination_actions,
                ),
            )
        if upload_kind == "unknown" and bio_data_file and bio_data_file.filename and not filename:
            bio_name = bio_data_file.filename or ""
            if not bio_name.lower().endswith((".xlsx", ".xlsm")):
                raise HTTPException(status_code=400, detail="Upload the CLITI Biodata file as .xlsx.")
            bio_content = await bio_data_file.read()
            bio_kind = _detect_cli_upload_kind(_load_cli_upload_rows(bio_content))
            if bio_kind != "cli_biodata":
                raise HTTPException(status_code=400, detail="The uploaded file is not recognized as CLITI Biodata.")
            bio_records, warnings = _parse_cli_biodata_workbook(bio_content)
            notice, warning_message, details, warnings = _apply_cli_biodata_records(session, bio_records, warnings, bio_name)
            notice_parts = [notice]
            warning_bits = [warning_message] if warning_message else []
            details = list(details)
            warnings = list(warnings)
            if nomination_notice:
                notice_parts.append(nomination_notice)
            details.extend(nomination_details)
            warnings.extend(nomination_warnings)
            if nomination_warning:
                warning_bits.append(nomination_warning)
            return templates.TemplateResponse(
                "cli.html",
                _cli_page_context(
                    request,
                    session,
                    grading_update_notice=" ".join(notice_parts),
                    grading_update_warning=" ".join(bit for bit in warning_bits if bit),
                    grading_update_details=details,
                    grading_warning_details=warnings,
                    nomination_mismatch_actions=nomination_actions,
                ),
            )
        if upload_kind == "cli_matrix":
            report_date = infer_report_date(filename) or date.today()
            summary_df = build_summary_df(content)
            overdue_df = build_sheet2_df(content)
            _save_cli_matrix_snapshots(session, report_date, summary_df, overdue_df)
            notice_parts = [f"CLI Matrix import complete for {report_date.strftime('%d-%m-%Y')}."]
            details: list[str] = []
            warnings: list[str] = []
            warning_bits = ["Detected a CLI Matrix file and routed it to the CLI Matrix snapshot importer."]
            if bio_data_file and bio_data_file.filename:
                bio_name = bio_data_file.filename or ""
                if not bio_name.lower().endswith((".xlsx", ".xlsm")):
                    raise HTTPException(status_code=400, detail="Upload the CLITI Biodata file as .xlsx.")
                bio_content = await bio_data_file.read()
                bio_kind = _detect_cli_upload_kind(_load_cli_upload_rows(bio_content))
                if bio_kind != "cli_biodata":
                    raise HTTPException(status_code=400, detail="The second file must be a CLITI Biodata workbook.")
                bio_records, warnings = _parse_cli_biodata_workbook(bio_content)
                bio_notice, bio_warning, details, warnings = _apply_cli_biodata_records(session, bio_records, warnings, bio_name)
                notice_parts.append(bio_notice)
                if bio_warning:
                    warning_bits.append(bio_warning)
            if nomination_notice:
                notice_parts.append(nomination_notice)
            details.extend(nomination_details)
            warnings.extend(nomination_warnings)
            if nomination_warning:
                warning_bits.append(nomination_warning)
            _save_li_grading_metadata(filename)
            return templates.TemplateResponse(
                "cli.html",
                _cli_page_context(
                    request,
                    session,
                    grading_update_notice=" ".join(notice_parts),
                    grading_update_warning=" ".join(warning_bits),
                    grading_update_details=details,
                    grading_warning_details=warnings,
                    nomination_mismatch_actions=nomination_actions,
                ),
            )

        if upload_kind == "cli_biodata":
            records, warnings = _parse_cli_biodata_workbook(content)
            _save_li_grading_metadata(filename)
            notice, warning_message, details, warnings = _apply_cli_biodata_records(session, records, warnings, filename)
            notice_parts = [notice]
            warning_bits = [warning_message] if warning_message else []
            details = list(details)
            warnings = list(warnings)
            if nomination_notice:
                notice_parts.append(nomination_notice)
            details.extend(nomination_details)
            warnings.extend(nomination_warnings)
            if nomination_warning:
                warning_bits.append(nomination_warning)
            return templates.TemplateResponse(
                "cli.html",
                _cli_page_context(
                    request,
                    session,
                    grading_update_notice=" ".join(bit for bit in notice_parts if bit),
                    grading_update_warning=" ".join(bit for bit in warning_bits if bit),
                    grading_update_details=details,
                    grading_warning_details=warnings,
                    nomination_mismatch_actions=nomination_actions,
                ),
            )

        records, warnings = _parse_li_grading_workbook(content)
        employees = session.exec(select(Employee)).all()
        canonical_by_id, alias_map, id_by_name = _build_cli_name_maps((employee.cli, employee.cli_id) for employee in employees)

        by_crew: dict[str, list[Employee]] = {}
        by_crew_name: dict[tuple[str, str], list[Employee]] = {}
        by_name_role: dict[tuple[str, str], list[Employee]] = {}
        for employee in employees:
            crew_key = (_clean_import_text(employee.crew_id) or "").upper()
            name_key = _normalize_import_name(employee.name)
            if crew_key:
                by_crew.setdefault(crew_key, []).append(employee)
                if name_key:
                    by_crew_name.setdefault((crew_key, name_key), []).append(employee)
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
            cli_name = _clean_import_text(record.get("cli_name"))
            cli_id = _clean_import_text(record.get("cli_id"))
            cli_name, cli_id = _canonicalize_cli_name(cli_name, cli_id, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)
            crew_key = str(record.get("crew_id") or "").upper()
            name_key = _normalize_import_name(record.get("name"))
            role_key = str(record.get("role") or "")
            target: Employee | None = None

            if crew_key and name_key:
                crew_name_matches = by_crew_name.get((crew_key, name_key), [])
                if len(crew_name_matches) == 1:
                    target = crew_name_matches[0]
                elif len(crew_name_matches) > 1:
                    filtered_matches = [
                        employee
                        for employee in crew_name_matches
                        if normalize_role(employee.role) == role_key
                    ]
                    if len(filtered_matches) == 1:
                        target = filtered_matches[0]
                    else:
                        warnings.append(f"CLI Grading {row_hint}: skipped because CREW ID + NAME matched multiple roster rows.")
                        skipped += 1
                        continue

            if target is None and crew_key:
                crew_matches = by_crew.get(crew_key, [])
                filtered_matches = [
                    employee
                    for employee in crew_matches
                    if _normalize_import_name(employee.name) == name_key and normalize_role(employee.role) == role_key
                ]
                if len(filtered_matches) == 1:
                    target = filtered_matches[0]
                elif len(filtered_matches) > 1:
                    warnings.append(f"CLI Grading {row_hint}: skipped because CREW ID, NAME, and DESIGNATION matched multiple roster rows.")
                    skipped += 1
                    continue
                elif crew_matches:
                    warnings.append(f"CLI Grading {row_hint}: skipped because CREW ID {crew_key} matched the roster but NAME / DESIGNATION did not match.")
                    skipped += 1
                    continue

            if target is None:
                if not name_key or not role_key:
                    warnings.append(f"CLI Grading {row_hint}: no matching CLI Roster row found.")
                    skipped += 1
                    continue
                fallback_matches = by_name_role.get((name_key, role_key), [])
                if len(fallback_matches) == 1:
                    target = fallback_matches[0]
                elif len(fallback_matches) > 1:
                    warnings.append(f"CLI Grading {row_hint}: skipped because NAME + DESIGNATION matched multiple CLI Roster rows.")
                    skipped += 1
                    continue
                else:
                    warnings.append(f"CLI Grading {row_hint}: no matching CLI Roster row found.")
                    skipped += 1
                    continue

            if target.id is not None and target.id in touched_ids:
                warnings.append(f"CLI Grading {row_hint}: skipped because that roster row already received a grading update from another row in this workbook.")
                skipped += 1
                continue

            new_grade = _clean_import_text(record.get("gradation"))
            new_due = record.get("grading_due")
            old_grade = _clean_import_text(target.gradation)
            old_due = target.grading_due
            old_cli, old_cli_id = _canonicalize_cli_name(target.cli, target.cli_id, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)

            if old_grade == new_grade and old_due == new_due and old_cli == cli_name and old_cli_id == cli_id:
                unchanged += 1
                if target.id is not None:
                    touched_ids.add(target.id)
                continue

            changes: list[str] = []
            if old_grade != new_grade:
                changes.append(f"Gradation: {_format_sync_value(old_grade)} -> {_format_sync_value(new_grade)}")
            if old_due != new_due:
                changes.append(f"Grading Due: {_format_sync_value(old_due)} -> {_format_sync_value(new_due)}")
            if old_cli != cli_name:
                changes.append(f"CLI: {_format_sync_value(old_cli)} -> {_format_sync_value(cli_name)}")
            if old_cli_id != cli_id:
                changes.append(f"CLI ID: {_format_sync_value(old_cli_id)} -> {_format_sync_value(cli_id)}")

            target.gradation = new_grade
            target.grading_due = new_due
            target.cli, target.cli_id = _canonicalize_cli_name(cli_name, cli_id, canonical_by_id=canonical_by_id, alias_map=alias_map, id_by_name=id_by_name)
            updated += 1
            if target.id is not None:
                touched_ids.add(target.id)
            details.append(
                f"Updated {target.name} ({target.crew_id or target.hrms or target.id}): " + "; ".join(changes)
            )

        session.commit()
        _normalize_employee_cli_names(session)
        _save_li_grading_metadata(filename)
        if updated == 0 and unchanged > 0 and skipped == 0:
            notice = "No change found in CLI grading file."
        else:
            notice_parts = []
            if updated:
                notice_parts.append(f"{updated} updated")
            if skipped:
                notice_parts.append(f"{skipped} skipped")
            if not notice_parts:
                notice_parts.append("No change found")
            notice = "CLI grading update complete: " + ", ".join(notice_parts) + "."
        warning_message = f"Mismatch / auto-fixed records: {len(warnings)}" if warnings else ""
        notice_parts = [notice]
        warning_bits = [warning_message] if warning_message else []
        if nomination_notice:
            notice_parts.append(nomination_notice)
        details.extend(nomination_details)
        warnings.extend(nomination_warnings)
        if nomination_warning:
            warning_bits.append(nomination_warning)
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_notice=" ".join(bit for bit in notice_parts if bit),
                grading_update_warning=" ".join(bit for bit in warning_bits if bit),
                grading_update_details=details,
                grading_warning_details=warnings,
                nomination_mismatch_actions=nomination_actions,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "CLI grading update failed."
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_error=detail,
            ),
            status_code=exc.status_code,
        )


@app.post("/cli/nomination-mismatch-apply")
async def apply_cli_nomination_mismatch(
    request: Request,
    existing_id: int = Form(...),
    incoming_payload: str = Form(...),
    action_key: str = Form(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        payload = json.loads(incoming_payload)
        employee = session.get(Employee, existing_id)
        if employee is None:
            raise HTTPException(status_code=404, detail="Existing employee row not found.")
        incoming_crew = _clean_import_text(payload.get("crew_id"))
        incoming_hrms = _clean_import_text(payload.get("hrms"))
        incoming_cli = _clean_import_text(payload.get("cli"))
        incoming_cli_id = _clean_import_text(payload.get("cli_id"))
        changes: list[str] = []
        if incoming_crew and _clean_import_text(employee.crew_id) != incoming_crew:
            changes.append(f"CREW ID: {_format_sync_value(employee.crew_id)} -> {_format_sync_value(incoming_crew)}")
            employee.crew_id = incoming_crew
        if incoming_hrms and _clean_import_text(employee.hrms) != incoming_hrms:
            changes.append(f"HRMS ID: {_format_sync_value(employee.hrms)} -> {_format_sync_value(incoming_hrms)}")
            employee.hrms = incoming_hrms
        old_cli, old_cli_id = _canonicalize_cli_name(employee.cli, employee.cli_id)
        new_cli, new_cli_id = _canonicalize_cli_name(incoming_cli, incoming_cli_id)
        if old_cli != new_cli:
            changes.append(f"CLI: {_format_sync_value(old_cli)} -> {_format_sync_value(new_cli)}")
        if old_cli_id != new_cli_id:
            changes.append(f"CLI ID: {_format_sync_value(old_cli_id)} -> {_format_sync_value(new_cli_id)}")
        employee.cli, employee.cli_id = _canonicalize_cli_name(new_cli, new_cli_id)
        session.add(employee)
        session.commit()
        _normalize_employee_cli_names(session)
        nomination_actions = _remove_cli_nomination_mismatch_action(action_key)
        notice = "CLI nomination mismatch applied to existing employee."
        details = [f"Updated {employee.name} ({employee.crew_id or employee.hrms or employee.id}): " + "; ".join(changes)] if changes else []
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_notice=notice,
                grading_update_details=details,
                nomination_mismatch_actions=nomination_actions,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "CLI nomination mismatch apply failed."
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=detail), status_code=exc.status_code)
    except Exception as exc:
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=str(exc)), status_code=500)


@app.post("/cli/nomination-mismatch-ignore")
async def ignore_cli_nomination_mismatch(
    request: Request,
    action_key: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        nomination_actions = _remove_cli_nomination_mismatch_action(action_key)
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_notice="CLI nomination mismatch dismissed.",
                nomination_mismatch_actions=nomination_actions,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "CLI nomination mismatch dismiss failed."
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=detail), status_code=exc.status_code)
    except Exception as exc:
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=str(exc)), status_code=500)


@app.post("/cli/nomination-mismatch-ignore-all")
async def ignore_all_cli_nomination_mismatches(
    request: Request,
    session: Session = Depends(get_session),
):
    try:
        nomination_actions = _clear_cli_nomination_mismatch_actions()
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_notice="All CLI nomination mismatches dismissed.",
                nomination_mismatch_actions=nomination_actions,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "CLI nomination mismatch dismiss failed."
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=detail), status_code=exc.status_code)
    except Exception as exc:
        return templates.TemplateResponse("cli.html", _cli_page_context(request, session, grading_update_error=str(exc)), status_code=500)

@app.post("/upload")
async def upload_employees(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    today = date.today()
    _sync_retired_employees(session, today)
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
    removed = 0
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
            if existing is None:
                pf_key = _emp_no_match_key(pf_no)
                if pf_key:
                    employees = session.exec(select(Employee)).all()
                    matches = [employee for employee in employees if _emp_no_match_key(employee.pf_no) == pf_key]
                    if len(matches) == 1:
                        existing = matches[0]
        if existing is None and hrms:
            existing = session.exec(select(Employee).where(Employee.hrms == hrms)).first()
        if existing is None:
            existing = session.exec(
                select(Employee).where(Employee.name == str(name).strip(), Employee.role == role)
            ).first()
        if retirement_date <= today:
            if existing:
                session.delete(existing)
                removed += 1
            continue
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
    if added == 0 and updated == 0 and removed == 0:
        raise HTTPException(status_code=400, detail="No rows imported. Check the sheet data or headers.")
    return RedirectResponse("/", status_code=303)


@app.post("/upload-seniority")
async def upload_seniority(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    _sync_retired_employees(session, date.today())
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
