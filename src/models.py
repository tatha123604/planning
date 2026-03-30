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
    crew_id: str | None = None
    dob: date | None = None
    doa: date | None = None
    do_report: date | None = None
    status: str | None = None
    working_at: str | None = None
    gradation: str | None = None
    grading_due: date | None = None
    cli: str | None = None
    cli_id: str | None = None
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


class SubNonContinuousSignOnSnapshot(SQLModel, table=True):
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


class SubNonContinuousSignOffSnapshot(SQLModel, table=True):
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


class CliDistributionTarget(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    cli_name: str
    cli_id: str | None = None
    active: bool = Field(default=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)


class CliDistributionPlan(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow, index=True)
    cli_count: int
    grade_a_total: int
    grade_b_total: int
    grade_c_total: int
    targets_json: str


class CliDistributionAssignment(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    plan_id: int = Field(index=True)
    employee_id: int | None = Field(default=None, index=True)
    name: str
    role: str
    gradation: str
    current_cli: str | None = None
    current_cli_id: str | None = None
    proposed_cli: str
    proposed_cli_id: str | None = None
