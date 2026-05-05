from __future__ import annotations

import pandas as pd

from services.calculator import process_punches
from services.exceptions import WorkException


def _base_input() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["De Carli Gonzalo", "De Carli Gonzalo"],
            "department": ["Produccion", "Produccion"],
            "schedule": ["Manana(07:30:00-12:00:00)", "Tarde(13:00:00-17:30:00)"],
            "work_date_raw": ["2026-04-01", "2026-04-01"],
            "entry_time_raw": ["07:30:56", "13:00:05"],
            "exit_time_raw": ["12:02:12", "17:32:45"],
        }
    )


def test_calculate_daily_and_monthly_hours_with_rounding() -> None:
    df = _base_input()

    daily, monthly, inconsistencies = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 480},
    )

    assert len(daily) == 1
    assert daily.iloc[0]["Estado"] == "Normal"
    assert int(daily.iloc[0]["Minutos reales"]) == 544
    assert int(daily.iloc[0]["Minutos redondeados"]) == 540
    assert int(daily.iloc[0]["Minutos extra"]) == 60
    assert daily.iloc[0]["Horas extra"] == "01:00"
    assert daily.iloc[0]["Horas totales"] == "09:00"
    assert "07:30 - 12:02" in daily.iloc[0]["Tramos trabajados"]

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Dias trabajados"]) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 540
    assert int(monthly.iloc[0]["Minutos extra"]) == 60
    assert monthly.iloc[0]["Horas extra"] == "01:00"
    assert monthly.iloc[0]["Horas totales"] == "09:00"

    assert inconsistencies.empty


def test_detects_required_inconsistencies_without_stopping_processing() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "", "30", "40"],
            "employee_name": ["Ana", "Bruno", "", "Carla"],
            "department": ["A", "B", "C", "D"],
            "schedule": ["", "", "", ""],
            "work_date_raw": ["2026-04-01", "fecha-mala", "2026-04-01", "2026-04-01"],
            "entry_time_raw": ["08:00", "08:00", "", "14:00"],
            "exit_time_raw": ["12:00", "17:00", "", "13:00"],
        }
    )

    daily, monthly, inconsistencies = process_punches(df)

    assert len(daily) == 1
    assert len(monthly) == 1
    assert not inconsistencies.empty

    issue_types = set(inconsistencies["Tipo de inconsistencia"])
    assert "ID de persona vacio" in issue_types
    assert "Nombre vacio" in issue_types
    assert "Fecha invalida" in issue_types
    assert "Ausente" in issue_types
    assert "Salida menor que entrada" in issue_types


def test_exception_replaces_missing_mark_inconsistency() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-01"],
            "entry_time_raw": [""],
            "exit_time_raw": [""],
        }
    )

    exceptions = [
        WorkException(
            employee_id="20",
            exception_date=pd.Timestamp("2026-05-01").date(),
            exception_type="Feriado",
            details="Dia del trabajador",
        )
    ]

    daily, monthly, inconsistencies = process_punches(df, exceptions=exceptions)

    assert len(daily) == 1
    assert len(monthly) == 1
    assert daily.iloc[0]["Estado"] == "Feriado"
    assert daily.iloc[0]["Horas totales"] == "00:00"
    assert daily.iloc[0]["Horas extra"] == "00:00"
    assert len(inconsistencies) == 1
    assert inconsistencies.iloc[0]["Tipo de inconsistencia"] == "Excepcion aplicada"
    assert "Feriado" in inconsistencies.iloc[0]["Detalle"]


def test_marks_late_and_absent_status_for_daily_cells() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20"],
            "employee_name": ["Ana", "Ana"],
            "department": ["A", "A"],
            "schedule": ["", ""],
            "work_date_raw": ["2026-05-01", "2026-05-04"],
            "entry_time_raw": ["08:10", ""],
            "exit_time_raw": ["12:00", ""],
            "late_raw": ["00:10", ""],
            "absent_raw": ["", "01:00"],
        }
    )

    daily, _, _ = process_punches(df, scheduled_minutes_by_employee={"20": 240})
    status_by_date = {str(row["Fecha"]): row["Estado"] for _, row in daily.iterrows()}

    assert status_by_date["2026-05-01"] == "Tarde"
    assert status_by_date["2026-05-04"] == "Ausente"


def test_present_day_never_classified_as_ausente_even_with_absent_marker() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20"],
            "employee_name": ["Ana"],
            "department": ["A"],
            "schedule": [""],
            "work_date_raw": ["2026-05-05"],
            "entry_time_raw": ["08:28"],
            "exit_time_raw": ["12:00"],
            "late_raw": ["00:58"],
            "absent_raw": ["04:00"],
        }
    )

    daily, _, _ = process_punches(df, scheduled_minutes_by_employee={"20": 240})

    assert len(daily) == 1
    assert daily.iloc[0]["Estado"] == "Tarde"


def test_late_is_determined_by_employee_schedule_start_minute() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "20", "20"],
            "employee_name": ["Ana", "Ana", "Ana", "Ana"],
            "department": ["A", "A", "A", "A"],
            "schedule": ["", "", "", ""],
            "work_date_raw": ["2026-05-06", "2026-05-06", "2026-05-07", "2026-05-07"],
            "entry_time_raw": ["07:30", "12:00", "07:31", "12:00"],
            "exit_time_raw": ["12:00", "16:30", "12:00", "16:30"],
            "late_raw": ["", "", "", ""],
            "absent_raw": ["", "", "", ""],
        }
    )

    daily, _, _ = process_punches(
        df,
        scheduled_minutes_by_employee={"20": 540},
        scheduled_start_minute_by_employee={"20": 450},
    )
    status_by_date = {str(row["Fecha"]): row["Estado"] for _, row in daily.iterrows()}

    assert status_by_date["2026-05-06"] == "Normal"  # 07:30 exacto
    assert status_by_date["2026-05-07"] == "Tarde"   # 07:31


def test_weekend_rule_only_counts_weekend_days_with_punches_and_as_overtime() -> None:
    df = pd.DataFrame(
        {
            "employee_id": ["20", "20", "20", "20", "20", "20"],
            "employee_name": ["Ana"] * 6,
            "department": ["A"] * 6,
            "schedule": [""] * 6,
            # 2026-05-08 viernes, 2026-05-09 sabado, 2026-05-10 domingo
            "work_date_raw": [
                "2026-05-08",
                "2026-05-08",
                "2026-05-09",
                "2026-05-09",
                "2026-05-10",
                "2026-05-10",
            ],
            "entry_time_raw": ["07:30", "13:30", "08:00", "", "08:00", ""],
            "exit_time_raw": ["12:00", "18:00", "12:00", "", "12:00", ""],
            "late_raw": ["", "", "", "", "", ""],
            "absent_raw": ["", "", "", "", "", ""],
        }
    )

    daily, monthly, _ = process_punches(df, scheduled_minutes_by_employee={"20": 540})

    # Viernes + Sabado con fichadas + Domingo con fichadas.
    assert len(daily) == 3

    friday = daily[daily["Fecha"].astype(str) == "2026-05-08"].iloc[0]
    saturday = daily[daily["Fecha"].astype(str) == "2026-05-09"].iloc[0]
    sunday = daily[daily["Fecha"].astype(str) == "2026-05-10"].iloc[0]

    assert int(friday["Minutos redondeados"]) == 540
    assert int(friday["Minutos extra"]) == 0

    assert int(saturday["Minutos redondeados"]) == 240
    assert int(saturday["Minutos extra"]) == 240
    assert saturday["Horas extra"] == "04:00"
    assert int(sunday["Minutos redondeados"]) == 240
    assert int(sunday["Minutos extra"]) == 240
    assert sunday["Horas extra"] == "04:00"

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Minutos extra"]) == 480


def test_empty_input_returns_empty_dataframes_with_expected_columns() -> None:
    df = pd.DataFrame(
        columns=[
            "employee_id",
            "employee_name",
            "department",
            "schedule",
            "work_date_raw",
            "entry_time_raw",
            "exit_time_raw",
        ]
    )

    daily, monthly, inconsistencies = process_punches(df)

    assert list(daily.columns) == [
        "ID de persona",
        "Nombre",
        "Fecha",
        "Departamento",
        "Estado",
        "Tramos trabajados",
        "Minutos reales",
        "Minutos redondeados",
        "Minutos extra",
        "Horas extra",
        "Horas totales",
    ]
    assert list(monthly.columns) == [
        "ID de persona",
        "Nombre",
        "Dias trabajados",
        "Minutos totales",
        "Minutos extra",
        "Horas extra",
        "Horas totales",
    ]
    assert list(inconsistencies.columns) == [
        "ID de persona",
        "Nombre",
        "Fecha",
        "Tipo de inconsistencia",
        "Detalle",
    ]
    assert daily.empty and monthly.empty and inconsistencies.empty
