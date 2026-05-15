from __future__ import annotations

import pandas as pd
import pytest

from services.exceptions import WorkException
from services.processor import (
    ValidationError,
    _filter_exceptions_to_report_month,
    _filter_to_report_month,
    _validate_results,
)


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


def _valid_daily_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ID de persona": ["20", "20"],
            "Nombre": ["Ana", "Ana"],
            "Fecha": [pd.Timestamp("2026-05-01").date(), pd.Timestamp("2026-05-02").date()],
            "Departamento": ["A", "A"],
            "Estado": ["Normal", "Tarde"],
            "Tramos trabajados": ["07:30 - 12:00", "07:31 - 12:01"],
            "Minutos reales": [270, 270],
            "Minutos redondeados": [270, 270],
            "Minutos extra": [0, 0],
            "Horas extra": ["00:00", "00:00"],
            "Horas totales": ["04:30", "04:30"],
        }
    )


def _valid_monthly_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ID de persona": ["20"],
            "Nombre": ["Ana"],
            "Dias trabajados": [2],
            "Minutos totales": [540],
            "Minutos extra": [0],
            "Horas extra": ["00:00"],
            "Horas totales": ["09:00"],
        }
    )


def test_validate_results_accepts_consistent_data() -> None:
    _validate_results(_valid_daily_df(), _valid_monthly_df())


def test_validate_results_rejects_duplicate_daily_employee_date() -> None:
    daily = _valid_daily_df()
    daily.loc[1, "Fecha"] = daily.loc[0, "Fecha"]
    monthly = pd.DataFrame(
        {
            "ID de persona": ["20"],
            "Nombre": ["Ana"],
            "Dias trabajados": [1],
            "Minutos totales": [540],
            "Minutos extra": [0],
            "Horas extra": ["00:00"],
            "Horas totales": ["09:00"],
        }
    )

    with pytest.raises(ValidationError, match="filas duplicadas por empleado/fecha"):
        _validate_results(daily, monthly)


def test_validate_results_rejects_daily_non_multiple_of_30() -> None:
    daily = _valid_daily_df()
    daily.loc[0, "Minutos redondeados"] = 275
    daily.loc[0, "Horas totales"] = "04:35"

    with pytest.raises(ValidationError, match="multiplo de 30"):
        _validate_results(daily, _valid_monthly_df())


def test_validate_results_rejects_daily_extra_greater_than_total() -> None:
    daily = _valid_daily_df()
    daily.loc[0, "Minutos extra"] = 300
    daily.loc[0, "Horas extra"] = "05:00"

    with pytest.raises(ValidationError, match="no puede superar Minutos redondeados"):
        _validate_results(daily, _valid_monthly_df())


def test_validate_results_rejects_ausente_with_worked_minutes() -> None:
    daily = _valid_daily_df()
    daily.loc[0, "Estado"] = "Ausente"

    with pytest.raises(ValidationError, match="no puede tener minutos trabajados"):
        _validate_results(daily, _valid_monthly_df())


def test_validate_results_rejects_monthly_overtime_mismatch() -> None:
    monthly = _valid_monthly_df()
    monthly.loc[0, "Minutos extra"] = 30
    monthly.loc[0, "Horas extra"] = "00:30"

    with pytest.raises(ValidationError, match="horas extra mensual"):
        _validate_results(_valid_daily_df(), monthly)
