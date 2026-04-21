from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta


@dataclass(frozen=True)
class PunchRecord:
    employee_id: str
    employee_name: str
    punch_datetime: datetime
    event_type: str | None = None


@dataclass(frozen=True)
class WorkSegment:
    employee_id: str
    employee_name: str
    work_date: date
    start_time: time
    end_time: time
    duration: timedelta


@dataclass(frozen=True)
class Inconsistency:
    employee_id: str
    employee_name: str
    work_date: date | None
    issue_type: str
    details: str
