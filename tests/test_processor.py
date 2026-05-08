from __future__ import annotations

import pandas as pd

from services.exceptions import WorkException
from services.processor import _filter_exceptions_to_report_month, _filter_to_report_month


def test_filter_to_report_month_keeps_predominant_month() -> None:
    df = pd.DataFrame(
        {
            "work_date_raw": [
                "2026-05-01",
                "2026-05-02",
                "2026-05-03",
                "2026-04-30",
            ],
            "employee_id": ["20", "20", "20", "20"],
        }
    )

    filtered, period = _filter_to_report_month(df)

    assert str(period) == "2026-05"
    assert len(filtered) == 3
    assert set(pd.to_datetime(filtered["work_date_raw"]).dt.to_period("M").astype(str)) == {"2026-05"}


def test_filter_exceptions_to_report_month() -> None:
    exceptions = [
        WorkException(employee_id="20", exception_date=pd.Timestamp("2026-05-01").date(), exception_type="Feriado", details=""),
        WorkException(employee_id="20", exception_date=pd.Timestamp("2026-04-30").date(), exception_type="Feriado", details=""),
    ]

    filtered = _filter_exceptions_to_report_month(exceptions, pd.Period("2026-05", freq="M"))

    assert len(filtered) == 1
    assert str(filtered[0].exception_date) == "2026-05-01"
