from __future__ import annotations

from datetime import date
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


class CliMatrixSummarySnapshot(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    report_date: date = Field(index=True)
    row_no: int = Field(index=True)
    cli_id: str
    cli_name: str
    alloted_desig: str
    fp_over_due: int
    oldest_fp_overdue_date: date | None = None


class CliMatrixOverdueSnapshot(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    report_date: date = Field(index=True)
    row_no: int = Field(index=True)
    cli_id: str
    cli_name: str
    alloted_desig: str
    fp_over_due: int
    oldest_fp_overdue_date: date | None = None
    counsel_over_due: int
    oldest_counsel_overdue_date: date | None = None
    grading_overdue: int
    oldest_grading_overdue_date: date | None = None
    total_over_due_cases: int


class NonContinuousSignOnSnapshot(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    report_date: date = Field(index=True)
    row_no: int = Field(index=True)
    crew_id: str | None = None
    crew_name: str | None = None
    desig: str | None = None
    station: str | None = None
    event_time: str | None = None
    sup_id: str | None = None
    entry_point: str | None = None
    train_no: str | None = None
    loco_no: str | None = None
    duty_type: str | None = None
    route_stn: str | None = None
    reason: str | None = None


class NonContinuousSignOffSnapshot(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    report_date: date = Field(index=True)
    row_no: int = Field(index=True)
    crew_id: str | None = None
    crew_name: str | None = None
    desig: str | None = None
    station: str | None = None
    event_time: str | None = None
    sup_id: str | None = None
    entry_point: str | None = None
    train_no: str | None = None
    loco_no: str | None = None
    duty_type: str | None = None
    route_stn: str | None = None
    reason: str | None = None
