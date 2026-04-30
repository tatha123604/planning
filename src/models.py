from __future__ import annotations

from datetime import date, datetime
from sqlmodel import Field, SQLModel


class Employee(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str
    role: str
    hire_date: date
    retirement_date: date | None = Field(default=None, nullable=True)
    promotion_role: str | None = None
    promotion_ready_date: date | None = None
    seniority_rank: int | None = None
    category: str | None = None
    pf_no: str | None = None
    hrms: str | None = None
    dob: date | None = None
    doa: date | None = None
    do_report: date | None = None
    status: str | None = None
    working_at: str | None = None
    gradation: str | None = None
    cli: str | None = None
    pme_due: date | None = None
    technical_due: date | None = None
    transportation_due: date | None = None


class Requirement(SQLModel, table=True):
    role: str = Field(primary_key=True)
    needed: int


class SstsSnapshotRun(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    observed_at: datetime = Field(index=True)
    source_count: int = 0
    fetch_status: str = "ok"
    fetch_error: str | None = None


class SstsDeviceSnapshot(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    run_id: int = Field(index=True)
    observed_at: datetime = Field(index=True)
    observed_day: date = Field(index=True)
    device_id: int = Field(index=True)
    name: str
    uniqueid: str | None = None
    phone: str | None = None
    contact: str | None = None
    lastupdate: datetime | None = Field(default=None, index=True)
    offline_minutes: int | None = None
    attributes: str | None = None
