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

    daily, monthly, inconsistencies = process_punches(df)

    assert len(daily) == 1
    assert int(daily.iloc[0]["Minutos reales"]) == 544
    assert int(daily.iloc[0]["Minutos redondeados"]) == 540
    assert daily.iloc[0]["Horas totales"] == "09:00"
    assert "07:30 - 12:02" in daily.iloc[0]["Tramos trabajados"]

    assert len(monthly) == 1
    assert int(monthly.iloc[0]["Dias trabajados"]) == 1
    assert int(monthly.iloc[0]["Minutos totales"]) == 540
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
    assert "Ambos vacios" in issue_types
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

    assert daily.empty
    assert monthly.empty
    assert len(inconsistencies) == 1
    assert inconsistencies.iloc[0]["Tipo de inconsistencia"] == "Excepcion aplicada"
    assert "Feriado" in inconsistencies.iloc[0]["Detalle"]


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
        "Tramos trabajados",
        "Minutos reales",
        "Minutos redondeados",
        "Horas totales",
    ]
    assert list(monthly.columns) == [
        "ID de persona",
        "Nombre",
        "Dias trabajados",
        "Minutos totales",
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
