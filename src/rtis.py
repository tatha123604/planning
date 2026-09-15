"""Division-wise RTIS imports and event browsing."""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime, timedelta
from hashlib import sha256
from io import BytesIO
import json
import math
from pathlib import Path
import posixpath
from urllib.parse import urlencode
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import RedirectResponse, Response
from fastapi.templating import Jinja2Templates
from sqlalchemy import Column, LargeBinary, UniqueConstraint, func
from sqlalchemy.exc import IntegrityError
from sqlmodel import Field, SQLModel, Session, select

from .db import get_session

DIVISIONS = ("SDAH", "HWH", "ASN", "MLDT")
FOCUS_EVENTS = ("H", "J", "K")
ANALYSIS_TYPES = {"passenger": "Passenger Train Analysis", "goods": "Goods Train Analysis"}
MAX_BYTES = 20 * 1024 * 1024
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HEADERS = ("Sr.No.", "Device Id", "Loco No.", "Latitude", "Longitude", "Station",
           "Event Time", "Event Type", "Speed", "Division Code", "Reporting Time",
           "Train Number", "Train Start Date")
router = APIRouter()
templates = Jinja2Templates(directory=str(Path(__file__).resolve().parent.parent / "templates"))


class RtisUpload(SQLModel, table=True):
    __table_args__ = (UniqueConstraint("division", "digest"),)
    id: int | None = Field(default=None, primary_key=True)
    division: str = Field(index=True)
    filename: str
    digest: str
    uploaded_at: datetime = Field(default_factory=datetime.utcnow)
    first_day: date
    last_day: date
    row_count: int
    added_count: int
    event_counts: str
    content: bytes = Field(sa_column=Column(LargeBinary, nullable=False))


class RtisEvent(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    upload_id: int = Field(foreign_key="rtisupload.id", index=True)
    fingerprint: str = Field(unique=True, index=True)
    division: str = Field(index=True)
    event_type: str = Field(index=True)
    event_time: datetime = Field(index=True)
    station: str = Field(index=True)
    loco: str
    train: str
    speed: float | None = None
    source: str


def train_type_filters():
    """User rule: digits only = passenger; letters plus digits = goods."""
    train = func.trim(RtisEvent.train)
    passenger = (train != "") & ~train.op("GLOB")("*[^0-9]*")
    goods = train.op("GLOB")("*[A-Za-z]*") & train.op("GLOB")("*[0-9]*")
    return passenger, goods


def _xml(archive: ZipFile, name: str):
    data = archive.read(name)
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ValueError("Unsupported XML declarations in workbook.")
    return ET.fromstring(data)


def _timestamp(value: str, epoch: datetime) -> datetime:
    try:
        result = datetime.fromisoformat(value.strip())
        if result.tzinfo is not None:
            raise ValueError("Use local event times without timezone offsets.")
        return result
    except ValueError:
        try:
            serial = float(value)
            # Excel's Windows date system contains a fictitious 1900 leap day.
            adjustment = 1 if epoch.year == 1899 and 0 < serial < 60 else 0
            return epoch + timedelta(days=serial + adjustment)
        except (ValueError, OverflowError):
            raise ValueError("Expected an ISO date/time or Excel date value.") from None


def parse_rtis(content: bytes, division: str) -> list[dict]:
    """Read cell values directly: RTIS exports may contain invalid Excel fills.

    Formula cells are rejected, not evaluated. Original bytes are kept unchanged.
    """
    if division not in DIVISIONS:
        raise ValueError("Choose SDAH, HWH, ASN or MLDT.")
    try:
        with ZipFile(BytesIO(content)) as archive:
            if sum(item.file_size for item in archive.infolist()) > 150 * 1024 * 1024:
                raise ValueError("Workbook expands beyond the 150 MB limit.")
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                shared = ["".join(node.itertext()) for node in _xml(archive, "xl/sharedStrings.xml")]
            workbook = _xml(archive, "xl/workbook.xml")
            properties = workbook.find("s:workbookPr", NS)
            epoch = datetime(1904, 1, 1) if properties is not None and properties.get("date1904") in ("1", "true") else datetime(1899, 12, 30)
            links = {node.get("Id"): node for node in _xml(archive, "xl/_rels/workbook.xml.rels")}
            records = []
            for sheet in workbook.findall("s:sheets/s:sheet", NS):
                link = links[sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
                if link.get("TargetMode") == "External":
                    raise ValueError("External worksheet links are unsupported.")
                target = link.get("Target", "")
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
                if not target.startswith("xl/worksheets/"):
                    continue
                header = None
                for row in _xml(archive, target).findall("s:sheetData/s:row", NS):
                    cells = {}
                    for cell in row:
                        column = "".join(c for c in cell.get("r", "") if c.isalpha())
                        if cell.find("s:f", NS) is not None:
                            raise ValueError(f"{sheet.get('name')} row {row.get('r')}: formulas are not accepted in RTIS data.")
                        value = cell.find("s:v", NS)
                        raw = value.text or "" if value is not None else ""
                        if cell.get("t") == "s":
                            raw = shared[int(raw)]
                        elif cell.get("t") == "inlineStr":
                            raw = "".join(cell.find("s:is", NS).itertext())
                        cells[column] = raw.strip()
                    if not any(cells.values()):
                        continue
                    if header is None:
                        header = cells
                        missing = set(HEADERS) - set(header.values())
                        if missing:
                            raise ValueError(f"{sheet.get('name')}: missing columns: {', '.join(sorted(missing))}.")
                        if len(set(header.values())) != len(header.values()):
                            raise ValueError("Duplicate column names in workbook.")
                        continue
                    source = {name: cells.get(col, "") for col, name in header.items() if name in HEADERS}
                    label = f"{sheet.get('name')} row {row.get('r')}"
                    if source["Division Code"].upper() != division:
                        raise ValueError(f"{label}: Division Code is '{source['Division Code']}', expected {division}.")
                    try:
                        event_time = _timestamp(source["Event Time"], epoch)
                        if source["Reporting Time"]:
                            _timestamp(source["Reporting Time"], epoch)
                        speed = float(source["Speed"]) if source["Speed"] else None
                        if speed is not None and (not math.isfinite(speed) or speed < 0):
                            raise ValueError("Speed must be a non-negative number.")
                    except ValueError as exc:
                        raise ValueError(f"{label}: {exc}") from None
                    event_type = source["Event Type"].upper()
                    if not event_type or not source["Loco No."]:
                        raise ValueError(f"{label}: Event Type and Loco No. are required.")
                    # Ignore export serial/report-delivery time when identifying the same event.
                    identity = [division, source["Device Id"], source["Loco No."], source["Station"],
                                event_time.isoformat(), event_type, source["Latitude"], source["Longitude"]]
                    records.append(dict(division=division, event_type=event_type, event_time=event_time,
                                        station=source["Station"], loco=source["Loco No."],
                                        train=source["Train Number"], speed=speed,
                                        source=json.dumps(source),
                                        fingerprint=sha256(json.dumps(identity).encode()).hexdigest()))
                    if len(records) > 100_000:
                        raise ValueError("Upload at most 100,000 events per file.")
            if not records:
                raise ValueError("No RTIS event rows found in the workbook.")
            return records
    except (BadZipFile, ET.ParseError, KeyError, IndexError, TypeError, AttributeError, RuntimeError, NotImplementedError):
        raise ValueError("Cannot read this workbook. Upload a valid RTIS .xlsx export.") from None


def import_rtis(session: Session, filename: str, content: bytes, division: str) -> str:
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("Upload an .xlsx file.")
    if len(content) > MAX_BYTES:
        raise ValueError("Each file must be 20 MB or smaller.")
    if division not in DIVISIONS:
        raise ValueError("Choose SDAH, HWH, ASN or MLDT.")
    digest = sha256(content).hexdigest()
    if session.exec(select(RtisUpload.id).where(RtisUpload.division == division, RtisUpload.digest == digest)).first():
        return "Already uploaded; no duplicate events added."
    records = parse_rtis(content, division)
    unique = {row["fingerprint"]: row for row in records}
    existing = set()
    keys = list(unique)
    for offset in range(0, len(keys), 500):
        existing.update(session.exec(select(RtisEvent.fingerprint).where(RtisEvent.fingerprint.in_(keys[offset:offset + 500]))).all())
    new = [row for key, row in unique.items() if key not in existing]
    upload = RtisUpload(division=division, filename=filename.replace("\\", "/").split("/")[-1], digest=digest,
                        first_day=min(row["event_time"].date() for row in records),
                        last_day=max(row["event_time"].date() for row in records),
                        row_count=len(records), added_count=len(new),
                        event_counts=json.dumps(Counter(row["event_type"] for row in records)), content=content)
    try:
        session.add(upload)
        session.flush()
        session.add_all([RtisEvent(upload_id=upload.id, **row) for row in new])
        session.commit()
    except IntegrityError:
        session.rollback()
        raise ValueError("An overlapping upload was saved at the same time. Retry this file.") from None
    return f"Saved {len(new):,} new events; {len(records) - len(new):,} duplicate rows skipped."


@router.get("/rtis")
def rtis_page(request: Request, division: str = "SDAH", day: str = "", event: str = "HJK",
              page: int = 1, analysis: str = "passenger", view: str = "classified",
              speed: str = "",
              session: Session = Depends(get_session)):
    if speed not in ("", "30", "40", "50"):
        raise HTTPException(400, "Choose All speeds, 30+, 40+ or 50+.")
    if analysis not in ANALYSIS_TYPES or view not in ("classified", "unclassified"):
        raise HTTPException(400, "Invalid analysis selection.")
    if division not in DIVISIONS or event not in (*FOCUS_EVENTS, "HJK"):
        raise HTTPException(400, "Invalid division or event filter.")
    try:
        selected_day = date.fromisoformat(day) if day else None
    except ValueError:
        raise HTTPException(400, "Invalid date.") from None
    filters = [RtisEvent.division == division, RtisEvent.event_type.in_(FOCUS_EVENTS)]
    if speed:
        filters.append(RtisEvent.speed >= int(speed))
    if selected_day:
        filters.extend([RtisEvent.event_time >= datetime.combine(selected_day, datetime.min.time()),
                        RtisEvent.event_time < datetime.combine(selected_day + timedelta(days=1), datetime.min.time())])
    passenger, goods = train_type_filters()
    unknown = ~(passenger | goods)
    unclassified = session.exec(select(func.count()).select_from(RtisEvent).where(
        *filters, unknown)).one()
    if view == "unclassified":
        filters.append(unknown)
    else:
        filters.append(passenger if analysis == "passenger" else goods)
    counts = dict(session.exec(select(RtisEvent.event_type, func.count()).where(*filters).group_by(RtisEvent.event_type)).all())
    if event != "HJK":
        filters.append(RtisEvent.event_type == event)
    total = session.exec(select(func.count()).select_from(RtisEvent).where(*filters)).one()
    pages = max(1, (total + 99) // 100)
    page = min(max(1, page), pages)
    rows = session.exec(select(RtisEvent).where(*filters).order_by(RtisEvent.event_time.desc(), RtisEvent.id.desc()).offset((page - 1) * 100).limit(100)).all()
    # Select metadata only, never all the stored workbooks while rendering the page.
    history = session.exec(select(RtisUpload.id, RtisUpload.filename, RtisUpload.uploaded_at,
                                  RtisUpload.first_day, RtisUpload.last_day, RtisUpload.row_count,
                                  RtisUpload.added_count).where(RtisUpload.division == division)
                           .order_by(RtisUpload.id.desc()).limit(50)).all()
    return templates.TemplateResponse(request=request, name="rtis.html", context={
        "active_page": "rtis", "divisions": DIVISIONS, "division": division, "day": day,
        "event": event, "counts": counts, "rows": rows, "total": total, "history": history,
        "page": page, "pages": pages,
        "analysis": analysis, "analysis_types": ANALYSIS_TYPES, "view": view,
        "unclassified": unclassified, "speed": speed,
        "division_query": urlencode({"day": day, "event": event, "analysis": analysis, "view": view, "speed": speed}),
        "switch_query": urlencode({"division": division, "day": day, "event": event, "speed": speed}),
        "query": urlencode({"division": division, "day": day, "event": event, "analysis": analysis, "view": view, "speed": speed}),
        "notice": request.query_params.get("notice", ""),
    })


@router.post("/rtis/upload")
def rtis_upload(division: str = Form(...), files: list[UploadFile] = File(...),
                analysis: str = Form("passenger"),
                session: Session = Depends(get_session)):
    if division not in DIVISIONS or analysis not in ANALYSIS_TYPES:
        raise HTTPException(400, "Choose a valid division.")
    if len(files) > 10:
        raise HTTPException(400, "Upload at most 10 files at a time.")
    messages = []
    for file in files:
        filename = (file.filename or "workbook.xlsx").replace("\\", "/").split("/")[-1]
        try:
            message = import_rtis(session, filename, file.file.read(MAX_BYTES + 1), division)
        except ValueError as exc:
            session.rollback()
            message = f"Not saved: {exc}"
        finally:
            file.file.close()
        messages.append(f"{filename}: {message}")
    return RedirectResponse("/rtis?" + urlencode({"division": division, "analysis": analysis, "notice": " | ".join(messages)}), status_code=303)


@router.get("/rtis/uploads/{upload_id}/download")
def rtis_download(upload_id: int, session: Session = Depends(get_session)):
    upload = session.get(RtisUpload, upload_id)
    if upload is None:
        raise HTTPException(404, "Upload not found.")
    return Response(content=upload.content,
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="RTIS_{upload.division}_{upload.id}.xlsx"'})
