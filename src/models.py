from __future__ import annotations

from datetime import date
from sqlmodel import Field, SQLModel


class Employee(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    name: str
    role: str
    hire_date: date
    retirement_date: date
    promotion_role: str | None = None
    promotion_ready_date: date | None = None
    seniority_rank: int | None = None


class Requirement(SQLModel, table=True):
    role: str = Field(primary_key=True)
    needed: int