from __future__ import annotations

from pathlib import Path

import pandas as pd

from services.schedule_info import load_schedule_profiles, load_scheduled_minutes


def test_load_scheduled_minutes_for_split_and_continuous_shift(tmp_path: Path) -> None:
    sample = pd.DataFrame(
        {
            "Id": ["20", "63"],
            "Nombre": ["De Carli Gonzalo", "Poncio Cristina"],
            "horario ingreso Mañana": ["07:30:00", "08:15:00"],
            "Horario salida Mañana": ["12:00:00", None],
            "Horario Ingreso Tarde": ["13:15:00", None],
            "Horario Salida Tarde": ["17:45:00", "16:14:00"],
            "Horario corrido": ["NO", "SI"],
        }
    )
    file_path = tmp_path / "info.xlsx"
    sample.to_excel(file_path, index=False)

    result = load_scheduled_minutes(file_path)

    assert result["20"] == 540
    assert result["63"] == 480


def test_load_scheduled_minutes_prefers_empleados_sheet_in_date_template(tmp_path: Path) -> None:
    config_df = pd.DataFrame({"Tipos de ausencia": ["VACACIONES"]})
    empleados_df = pd.DataFrame(
        {
            "Id": ["20"],
            "Nombre": ["Ana"],
            "horario ingreso Mañana": ["07:30:00"],
            "Horario salida Mañana": ["12:00:00"],
            "Horario Ingreso Tarde": ["13:00:00"],
            "Horario Salida Tarde": ["17:30:00"],
            "Horario corrido": ["NO"],
        }
    )

    file_path = tmp_path / "date.xlsx"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        config_df.to_excel(writer, sheet_name="Config", index=False)
        empleados_df.to_excel(writer, sheet_name="Empleados", index=False)

    result = load_scheduled_minutes(file_path)

    assert result["20"] == 540


def test_load_schedule_profiles_parses_custom_working_days(tmp_path: Path) -> None:
    sample = pd.DataFrame(
        {
            "Id": ["118"],
            "Nombre": ["Zeller Lucas Ezequiel"],
            "horario ingreso Mañana": ["07:30:00"],
            "Horario salida Mañana": ["12:00:00"],
            "Horario Ingreso Tarde": [None],
            "Horario Salida Tarde": [None],
            "Horario corrido": ["SI"],
            "Dias": ["lunes miercoles y viernes"],
        }
    )
    file_path = tmp_path / "date.xlsx"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        sample.to_excel(writer, sheet_name="Empleados", index=False)

    profiles = load_schedule_profiles(file_path)

    assert profiles["118"]["scheduled_minutes"] == 270
    assert set(profiles["118"]["working_weekdays"]) == {0, 2, 4}
