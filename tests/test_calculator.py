from datetime import datetime

import pandas as pd

from services.calculator import process_punches


def test_calculate_daily_and_monthly_hours() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["100", "100", "100", "100"],
            "employee_name": ["Ana", "Ana", "Ana", "Ana"],
            "punch_datetime": [
                datetime(2026, 3, 3, 8, 0),
                datetime(2026, 3, 3, 12, 0),
                datetime(2026, 3, 3, 13, 0),
                datetime(2026, 3, 3, 17, 0),
            ],
            "work_date": [datetime(2026, 3, 3).date()] * 4,
        }
    )

    daily, monthly, inconsistencies = process_punches(df)

    assert len(daily) == 1
    assert daily.iloc[0]["total_hours"] == "08:00"
    assert int(daily.iloc[0]["total_minutes"]) == 480

    assert len(monthly) == 1
    assert monthly.iloc[0]["monthly_total_hours"] == "08:00"
    assert inconsistencies.empty


def test_detects_odd_punch_count() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["100", "100", "100"],
            "employee_name": ["Ana", "Ana", "Ana"],
            "punch_datetime": [
                datetime(2026, 3, 3, 8, 0),
                datetime(2026, 3, 3, 12, 0),
                datetime(2026, 3, 3, 13, 0),
            ],
            "work_date": [datetime(2026, 3, 3).date()] * 3,
        }
    )

    _, _, inconsistencies = process_punches(df)

    assert not inconsistencies.empty
    assert "Cantidad impar" in set(inconsistencies["issue_type"])
