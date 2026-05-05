from __future__ import annotations

from pathlib import Path

import pandas as pd

from services.schedule_info import load_scheduled_minutes


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
