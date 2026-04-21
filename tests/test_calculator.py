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
codex/create-desktop-app-for-attendance-processing-18do52
    assert daily.iloc[0]["Horas totales"] == "08:00"
    assert int(daily.iloc[0]["Minutos totales"]) == 480
    assert "Ingreso 08:00" in daily.iloc[0]["Marcaciones (I/S)"]
    assert "Salida 12:00" in daily.iloc[0]["Marcaciones (I/S)"]

    assert len(monthly) == 1
    assert monthly.iloc[0]["Horas mensuales"] == "08:00"

    assert daily.iloc[0]["total_hours"] == "08:00"
    assert int(daily.iloc[0]["total_minutes"]) == 480

    assert len(monthly) == 1
    assert monthly.iloc[0]["monthly_total_hours"] == "08:00"
 main
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
codex/create-desktop-app-for-attendance-processing-18do52
    assert "Cantidad impar" in set(inconsistencies["Tipo de inconsistencia"])


def test_empty_input_returns_empty_dataframes_with_expected_columns() -> None:
    df = pd.DataFrame(columns=["employee_id", "employee_name", "punch_datetime", "work_date"])

    daily, monthly, inconsistencies = process_punches(df)

    assert list(daily.columns) == [
        "Legajo",
        "Nombre",
        "Fecha",
        "Cantidad de fichadas",
        "Marcaciones (I/S)",
        "Tramos trabajados",
        "Horas totales",
        "Minutos totales",
    ]
    assert list(monthly.columns) == ["Legajo", "Nombre", "Minutos totales", "Horas mensuales"]
    assert list(inconsistencies.columns) == ["Legajo", "Nombre", "Fecha", "Tipo de inconsistencia", "Detalle"]
    assert daily.empty and monthly.empty and inconsistencies.empty

    assert "Cantidad impar" in set(inconsistencies["issue_type"])
main
