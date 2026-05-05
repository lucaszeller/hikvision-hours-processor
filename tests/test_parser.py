from __future__ import annotations

from datetime import datetime

import pandas as pd

from services.parser import detect_column_mapping


def test_detect_column_mapping_hikvision_spanish_headers() -> None:
    columns = [
        "Indice",
        "ID de persona",
        "Nombre",
        "Departamento",
        "Posicion",
        "Fecha",
        "Horario",
        "Registro de entrada",
        "Registro de salida",
    ]

    mapping = detect_column_mapping(columns)

    assert mapping["employee_id"] == "ID de persona"
    assert mapping["employee_name"] == "Nombre"
    assert mapping["work_date"] == "Fecha"
    assert mapping["entry_time"] == "Registro de entrada"
    assert mapping["exit_time"] == "Registro de salida"
    assert mapping["department"] == "Departamento"


def test_date_and_time_strings_parse_with_pandas() -> None:
    date_part = pd.Series([datetime(2026, 4, 1), datetime(2026, 4, 2)])
    time_part = pd.Series(["07:30:00", "12:02:12"])

    joined = pd.to_datetime(date_part.dt.strftime("%Y-%m-%d") + " " + time_part, errors="coerce")

    assert joined.isna().sum() == 0
    assert joined.iloc[0].hour == 7
    assert joined.iloc[1].minute == 2
